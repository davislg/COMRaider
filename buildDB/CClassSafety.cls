VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClassSafety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Option Explicit

Public clsid As String
Public InProcServer As String
Public InProcHandler As String
Public LocalServer As String
Public typelib As String

Public isSafe As Boolean
Public wasRegistered As Boolean

Public regSafeForScript As Boolean
Public regSafeForInit As Boolean
Public HasIObjSafety As Boolean

Public IDispSafe As Long
Public IDispExSafe  As Long
Public IPersistSafe As Long
Public IPSteamSafe  As Long
Public IPStorageSafe  As Long

Friend Sub SetGeneralFlag()

    If HasIObjSafety Or regSafeForScript Or regSafeForInit Then isSafe = True
    
    'On Error Resume Next
    
    If isSafe Then
            Dim progid As String, inproc As String, desc As String, handler As String
            Dim server As String, typelib As String
            
            desc = reg.ReadValue("\CLSID\" & clsid, "")
            progid = reg.ReadValue("\CLSID\" & clsid & "\ProgID", "")
            inproc = reg.ReadValue("\CLSID\" & clsid & "\InProcServer32", "")
            handler = reg.ReadValue("\CLSID\" & clsid & "\InProcHandler32", "")
            server = reg.ReadValue("\CLSID\" & clsid & "\Localserver32", "")
            typelib = reg.ReadValue("\CLSID\" & clsid & "\TypeLib", "")
            
            Const fields = "clsid,description,ProgID,InProcHandler32," & _
                           "Localserver32,InProcServer32,sDate,safeforscript," & _
                           "safeforinit,hasobjsafety,typelib,idispsafe," & _
                           "idispexsafe,ipersistsafe,ipstreamsafe,ipstoragesafe"
             
            Dim cnt As Long, sql As String
            
            cnt = cn.Execute("Select count(autoid) as cnt from tblGUIDs where clsid='" & clsid & "'")!cnt
            
            If cnt < 1 Then
            
                Insert "tblGUIDs", fields, _
                       clsid, desc, progid, handler, _
                       server, inproc, Format(Now, "m.d.yy"), _
                       IIf(regSafeForScript, 1, 0), IIf(regSafeForInit, 1, 0), _
                       IIf(HasIObjSafety, 1, 0), typelib, IDispSafe, IDispExSafe, _
                       IPersistSafe, IPSteamSafe, IPStorageSafe
                   
            Else
                
                'in case not filled out yet by quick scan method we update now
                sql = "Update tblGUIDs set hasobjsafety=" & IIf(HasIObjSafety, 1, 0) & _
                       " , idispsafe=" & IDispSafe & _
                       " , idispexsafe=" & IDispExSafe & _
                       " , ipersistsafe=" & IPersistSafe & _
                       " , ipstreamsafe=" & IPSteamSafe & _
                       " , ipstoragesafe=" & IPStorageSafe & _
                       " where clsid='" & clsid & "'"
                       
                cn.Execute sql
                
            End If
                       
                    
    End If
    
    
End Sub

Private Sub Insert(tblName, fields, ParamArray params())
    Dim sql As String, i As Integer, values(), tn As String
    
    values() = params() 'force byval
    
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Then
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next

    sql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    sql = Replace(sql, "____", Join(values, ","))
    cn.Execute sql
    
End Sub

Function GetReport() As String
    Dim tmp() As String
    
    push tmp, "Report for Clsid: " & clsid
    
    If wasRegistered Then
    
        push tmp, "RegKey Safe for Script: " & regSafeForScript
        push tmp, "RegKey Safe for Init: " & regSafeForInit
        push tmp, "Implements IObjectSafety: " & HasIObjSafety
        
        If IDispSafe > 0 Then push tmp, "IDisp Safe: " & FlagToText(IDispSafe)
        If IDispExSafe > 0 Then push tmp, "IDispEx Safe: " & FlagToText(IDispExSafe)
        If IPersistSafe > 0 Then push tmp, "IPersist Safe: " & FlagToText(IPersistSafe)
        If IPSteamSafe > 0 Then push tmp, "IPSteam Safe: " & FlagToText(IPSteamSafe)
        If IPStorageSafe > 0 Then push tmp, "IPStorage Safe: " & FlagToText(IPStorageSafe)
        
    Else
        push tmp, "NOT REGISTERED ON YOUR MACHINE"
    End If
    
    GetReport = Join(tmp, vbCrLf)
    
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


Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
