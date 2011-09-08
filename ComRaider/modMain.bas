Attribute VB_Name = "modMain"
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Global fso As New CFileSystem2
Global dlg As New clsCmnDlg
Global reg As New clsRegistry2
Global cnDistro As New Connection
Global cn As New Connection
Global killbitted As New Collection

Global Const AuditColor As Long = &H66CC22

Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Public Type udtOptions
    DistributedMode As Boolean
    DistributedModeLastAutoid As Long
    ExternalDebugger As String
    UserName As String
    AllowObjArgs As Boolean
    OnlyDefInf As Boolean
    CfgFileVersion As Long
    UseSimpleRegScanner As Boolean
    UseApiLogger As Boolean
    ApiFilters As String
    IEPath As String
    ApiTriggers As String
    UseSymbols As Boolean
    SymPath As String
End Type

Public Options As udtOptions

Function QueryDsn(sql As String) As Recordset
    On Error GoTo hell
    If cnDistro.State = 0 Then cnDistro.Open
    Set QueryDsn = cnDistro.Execute(sql)
    
    Exit Function
hell:
    MsgBox "Error in QueryDsn: " & Err.Description, vbExclamation
    On Error Resume Next
    Set QueryDsn = cn.Execute(sql) 'query local access db to get empty recordset so no crashs
End Function

Sub DsnInsert(tblName, fields, ParamArray params())
    On Error GoTo hell
    
    Dim sql As String, i As Integer, values(), tn As String
    
    values() = params() 'force byval
    
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "field" Then
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next

    sql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    sql = Replace(sql, "____", Join(values, ","))
    If cnDistro.State = 0 Then cnDistro.Open
    cnDistro.Execute sql
    
    Exit Sub
hell:
    MsgBox "Error in DsnInert: " & Err.Description, vbExclamation
End Sub

Sub ShowObjSafetyReport(clsid As String)
    Dim l As String
    
    l = App.path & IIf(IsIde, "\..", "") & "\builddb.exe"
    If Not fso.FileExists(l) Then
        MsgBox "Could not find: " & l, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Shell """" & l & """ """ & clsid & """ /report", vbNormalFocus
End Sub

Function GetSavedSafetyReport(clsid As String, report As String) As Boolean
    
    Dim rs As Recordset
    Dim tmp() As String
    Dim IDispSafe As Long
    Dim IDispExSafe  As Long
    Dim IPersistSafe As Long
    Dim IPSteamSafe  As Long
    Dim IPStorageSafe  As Long

    On Error Resume Next
    
    Set rs = cn.Execute("Select * from tblguids where clsid='" & clsid & "'")
    
    report = Empty
    If rs.EOF Then Exit Function
    
    If Not IsNull(rs!safeForScript) Then
        push tmp, "RegKey Safe for Script: " & CBool(rs!safeForScript)
    End If
    
    If Not IsNull(rs!safeForInit) Then
        push tmp, "RegKey Safe for Init: " & CBool(rs!safeForInit)
    End If
    
    If Not IsNull(rs!hasobjSafety) Then
        push tmp, "Implements IObjectSafety: " & CBool(rs!hasobjSafety)
    End If
    
    'these could be null in which case we want them to be unified to 0 first
    IDispSafe = CLng(rs!IDispSafe)
    IDispExSafe = CLng(rs!IDispExSafe)
    IPersistSafe = CLng(rs!IPersistSafe)
    IPSteamSafe = CLng(rs!ipstreamsafe)
    IPStorageSafe = CLng(rs!IPStorageSafe)
    
    If IDispSafe > 0 Then push tmp, "IDisp Safe: " & FlagToText(IDispSafe)
    If IDispExSafe > 0 Then push tmp, "IDispEx Safe: " & FlagToText(IDispExSafe)
    If IPersistSafe > 0 Then push tmp, "IPersist Safe: " & FlagToText(IPersistSafe)
    If IPSteamSafe > 0 Then push tmp, "IPStream Safe: " & FlagToText(IPSteamSafe)
    If IPStorageSafe > 0 Then push tmp, "IPStorage Safe: " & FlagToText(IPStorageSafe)

    report = Join(tmp, vbCrLf)
    GetSavedSafetyReport = True
    
End Function

Private Function FlagToText(flag As Long) As String
    
    'Public Const INTERFACESAFE_FOR_UNTRUSTED_CALLER = 1 'Caller of interface may be untrusted
    'Public Const INTERFACESAFE_FOR_UNTRUSTED_DATA = 2   'Data passed into interface may be untrusted
    'Public Const INTERFACE_USES_DISPEX = 4                'Object knows to use IDispatchEx")
    'Public Const INTERFACE_USES_SECURITY_MANAGER = 8      'Object knows to use IInternetHostSecurityManager

    Dim ret As String, extras As String
    
    If (flag Or 1) = flag Then ret = "caller"
    If (flag Or 2) = flag Then ret = ret & IIf(Len(ret) > 0, ",", "") & "data"
    If (flag Or 4) = flag Then extras = "USES_IDISPEX"
    If (flag Or 8) = flag Then extras = extras & IIf(Len(extras) > 0, ",", "") & "USES_SEC_MGR"
    
    FlagToText = IIf(Len(ret) > 0, " Safe for untrusted: " & ret, "") & "  " & extras
    
End Function


Function StripQuotes(ByVal x)
    x = Replace(x, "'", Empty)
    StripQuotes = Replace(x, """", Empty)
End Function

Function ExpandPath(ByVal fPath As String) As String
    Dim x As Long
    Dim tmp As String
    
    On Error Resume Next
    
    fPath = StripQuotes(fPath)
    x = InStrRev(fPath, "%")
    If x > 0 Then
        env = Mid(fPath, 1, x)
        fPath = Replace(fPath, env, Environ(Replace(env, "%", "")))
    End If
        
    If InStr(LCase(fPath), ":\") < 1 Then
        tmp = Environ("WinDIR") & "\" & fPath
        If fso.FileExists(tmp) Then
            fPath = tmp
        Else
            tmp = Environ("WinDIR") & "\System32\" & fPath
            If fso.FileExists(tmp) Then fPath = tmp
        End If
    End If
    
    ExpandPath = fPath
    
End Function

Sub LoadOptions()
    Dim ff As String
    Dim f As Long
    Dim tmp() As String, t, v
        
    On Error Resume Next
    
    Const CFG_FILE_VER As Integer = 6
    
    ff = App.path & "\options.dat"
    If fso.FileExists(ff) Then
        f = FreeFile
        Open ff For Binary As f
        Get f, , Options
        Close f
    Else
        With Options
            .DistributedMode = True
            .UserName = "UserX"
            .AllowObjArgs = False
            .OnlyDefInf = True
            .UseSimpleRegScanner = True
            .UseApiLogger = True
            .CfgFileVersion = CFG_FILE_VER
            .IEPath = "C:\Program Files\Internet Explorer\iexplore.exe"
            .SymPath = Environ("_NT_SYMBOL_PATH")
            If Len(Options.SymPath) > 0 Then Options.UseSymbols = True
        End With
    End If
    
    If Options.CfgFileVersion < CFG_FILE_VER Then
        Options.CfgFileVersion = CFG_FILE_VER
        Options.UseSimpleRegScanner = True
        Options.UseApiLogger = True
        Options.ApiFilters = "Kaspersky,Symantec"
        Options.IEPath = "C:\Program Files\Internet Explorer\iexplore.exe"
        Options.ApiTriggers = Empty
        Options.SymPath = Environ("_NT_SYMBOL_PATH")
        If Len(Options.SymPath) > 0 Then Options.UseSymbols = True
    End If
    
    cnDistro.ConnectionString = "DSN=COMRaider;"
    cnDistro.Close
    Err.Clear
    
    With MDIForm1
        If Options.DistributedMode Then
                .Visible = True
                .lblDsn.Visible = True
                .lblDsn.Refresh
                cnDistro.Open
                If Err.Number <> 0 Then Options.DistributedMode = False 'no connection
                If Options.DistributedMode Then
                        .tmrKeepalive.Enabled = True
                        .mnuAuditLogs.Enabled = True
                        .mnuDistNotes.Enabled = True
                        .mnuUploadOfflineAudits.Enabled = True
                End If
                .lblDsn.Visible = False
                cnDistro.Close
        End If
    End With
    
    
    reg.hive = HKEY_LOCAL_MACHINE
    Const base = "\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility"
    tmp() = reg.EnumKeys(base)
    
    For Each t In tmp
        v = reg.ReadValue(base & "\" & t, "Compatibility Flags")
        If v = &H400 Then killbitted.Add t, t
    Next
            
End Sub
    
Sub SaveOptions()
    Dim ff As String
    Dim f As Long
    On Error Resume Next
    ff = App.path & "\options.dat"
    f = FreeFile
    Open ff For Binary As f
    Put f, , Options
    Close f
End Sub



Sub dbg(msg)
    'debug.Print msg
End Sub

Public Function UserDeskTopFolder() As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, 0, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      UserDeskTopFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
  
End Function


Function GetProgID(GUID As String) As String
    Dim tmp As String
    Dim f As String
    
    reg.hive = HKEY_CLASSES_ROOT
    If Len(GUID) = 0 Then Exit Function
    
    f = "\CLSID\" & f
        
    If reg.keyExists(f) Then
        f = f & "\ProgID"
        If reg.keyExists(f) Then
            f = reg.ReadValue(f, "")
            GetProgID = f
            Exit Function
        End If
    End If
    
    tmp = Split(GUID, "-")(0)
    GetProgID = Right(tmp, Len(tmp) - 1)

End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Sub glue(ary, value) 'this modifies parent ary object
    On Error GoTo hell
    ary(UBound(ary)) = ary(UBound(ary)) & " " & value
Exit Sub
hell: push ary, value
      Stop
End Sub



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


Sub SizeLV(lv As ListView, Optional first As Boolean = False)
    Dim c As Long
    With lv
        If first Then
            c = .Width
            For i = 2 To .ColumnHeaders.Count
                c = c - .ColumnHeaders(i).Width
            Next
            .ColumnHeaders(1).Width = c - 200
        Else
            c = .ColumnHeaders.Count
            .ColumnHeaders(c).Width = .Width - lv.ColumnHeaders(c).Left - 300
        End If
    End With
End Sub

Function SafeFreeFileName(pFolder As String, ext As String) As String
    On Error GoTo hell
    Dim i As Integer
    Dim tmp As String
top:
    tmp = fso.GetFreeFileName(pFolder, ext)
    SafeFreeFileName = tmp
    
    Exit Function
hell:
      i = i + 1
      If i > 5 Then Exit Function
      GoTo top
    
End Function


Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    IsIde = False
    Exit Function
hell: IsIde = True
End Function



Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    If IsObject(c(val)) Then
        Set t = c(val)
    Else
        t = c(val)
    End If
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

'call backs have to be in public modules
Sub ModuleListCallBack(ByVal pName As Long, ByVal base As Long, ByVal size As Long)
    Dim buf() As Byte
    Dim name As String
    Dim n As Long
    
    ReDim buf(255)
    
    CopyMemory buf(0), ByVal pName, 255
    name = StrConv(buf, vbUnicode)
    n = InStr(name, Chr(0))
    name = Mid(name, 1, n - 1)
    
    'MsgBox "Name" & name & " Base" & Hex(base) & " Size" & Hex(size)
    frmCrashMon.Crash.AddModule name, base, size
    
End Sub

Sub StackWalkCallBack(ByVal frame As Long, ByVal eip As Long, ByVal retAddr As Long, ByVal frameptr As Long, ByVal stackPtr As Long)
    
    If retAddr > 0 Then
        frmCrashMon.Crash.AddStackFrame frame, eip, retAddr, frameptr, stackPtr
    End If

End Sub

Function ExtractGUID(WSFFile As String) As String

    If Not fso.FileExists(WSFFile) Then Exit Function
    
    Dim a As Long, b As Long, tmp As String
    
    Const marker1 = "classid='clsid:"
    
    tmp = fso.ReadFile(WSFFile)
    a = InStr(tmp, marker1)
    If a > 0 Then
        a = a + Len(marker1)
        b = InStr(a, tmp, "'")
        If b > a Then
            tmp = Mid(tmp, a, b - a)
            ExtractGUID = "{" & tmp & "}"
        End If
    End If
    
End Function

Function EditAuditNote(clsid As String)
    Dim Data As String, sql As String, cnt As String
    Dim rs As Recordset
    
    On Error Resume Next
    
    cnt = QueryDsn("Select count(autoid) as cnt from tblauditlog where clsid='" & clsid & "'")!cnt
    
    If cnt = 0 Then
        MsgBox "Not Audit Events for this clsid", vbInformation
        Exit Function
    End If
    
    If cnt = 1 Then
        Set rs = QueryDsn("Select * from tblauditlog where clsid='" & clsid & "'")
        Data = IIf(IsNull(rs!notes), "", rs!notes)
        Data = frmMsg.GetData("Enter Audit Notes", Data)
        sql = "Update tblauditlog set notes='" & Replace(Data, "'", "") & "' where autoid=" & rs!autoid
        QueryDsn sql
    Else
        frmEditAuditNotes.EditNotesFor clsid
    End If
        
        
End Function

Function LikeAnyOfThese(ByVal sIn, ByVal sCmp) As Boolean
    Dim tmp() As String, i As Integer
    On Error GoTo hell
    sIn = LCase(sIn)
    sCmp = LCase(sCmp)
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = "*" & Trim(tmp(i)) & "*"
        If Len(tmp(i)) > 0 And sIn Like tmp(i) Then
            LikeAnyOfThese = True
            Exit Function
        End If
    Next
hell:
End Function

Function AnyOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    On Error GoTo hell
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
hell:
End Function

