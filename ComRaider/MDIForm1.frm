VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "ComRaider"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   10650
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   0
         Picture         =   "MDIForm1.frx":030A
         ScaleHeight     =   1635
         ScaleWidth      =   10695
         TabIndex        =   1
         Top             =   -180
         Width           =   10695
         Begin VB.Timer tmrKeepalive 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   5040
            Top             =   840
         End
         Begin VB.Timer tmrCheese 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   6720
            Top             =   840
         End
         Begin VB.PictureBox pictStartOn 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   7200
            MouseIcon       =   "MDIForm1.frx":AB01
            MousePointer    =   99  'Custom
            Picture         =   "MDIForm1.frx":AE0B
            ScaleHeight     =   510
            ScaleWidth      =   1560
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.PictureBox pictStartOff 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   9120
            Picture         =   "MDIForm1.frx":E58F
            ScaleHeight     =   510
            ScaleWidth      =   1560
            TabIndex        =   2
            Top             =   480
            Width           =   1560
         End
         Begin MSWinsockLib.Winsock ws 
            Index           =   0
            Left            =   6300
            Top             =   840
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Label lblDsn 
            BackStyle       =   0  'Transparent
            Caption         =   "... Trying to connect to Distributed Server"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            MouseIcon       =   "MDIForm1.frx":11D13
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1080
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label lblView 
            BackStyle       =   0  'Transparent
            Caption         =   "View"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            MouseIcon       =   "MDIForm1.frx":1201D
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1080
            Width           =   435
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10680
      TabIndex        =   5
      Top             =   0
      Width           =   10680
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuEditVbs 
         Caption         =   "Edit BuildArgs.vbs"
      End
      Begin VB.Menu mnuAuditLogs 
         Caption         =   "Audit Logs"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDistNotes 
         Caption         =   "Distributed Notes"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUploadOfflineAudits 
         Caption         =   "Upload Offline Audits"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpnLinks 
         Caption         =   "Help && Links"
         Begin VB.Menu mnuHelp 
            Caption         =   "Help File"
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "About COMRaider"
         End
         Begin VB.Menu mnuVCPLink 
            Caption         =   "iDefense VCP Program"
         End
         Begin VB.Menu mnuLabsSoftware 
            Caption         =   "iDefense Labs Software"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         disassembler functionality provided by olly.dll which
'         is a modified version of the OllyDbg GPL source from
'         Oleh Yuschuk Copyright (C) 2001 - http://ollydbg.de
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


Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private setbg As Boolean
Private txtGUID As String
Private webPage As String
Private ticks As Long




 

Private Sub lblView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     PopupMenu mnuPopup
End Sub

 

Private Sub mnuAbout_Click()
    On Error Resume Next
    frmAbout.Show 1, Me
End Sub

Private Sub mnuAuditLogs_Click()
    frmAuditLogs.Visible = True
End Sub

Private Sub mnuDistNotes_Click()
    frmDistNotes.Visible = True
End Sub

Private Sub mnuEditVbs_Click()
    On Error Resume Next
    Dim p As String
    If IsIde Then p = "\.."
    p = App.path & p & "\buildargs.vbs"
    Shell "notepad '" & p & "'", vbNormalFocus
End Sub

Private Sub mnuHelp_Click()
    
    Dim pth As String
    pth = App.path & IIf(IsIde, "\..", "") & "\Comraider.chm"
    
    If Not fso.FileExists(pth) Then
        MsgBox "File not found: " & pth, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "hh.exe """ & pth & """", vbNormalFocus
    
End Sub

Private Sub mnuLabsSoftware_Click()
    On Error Resume Next
    ShellExecute 0, "Open", "http://labs.idefense.com/labs-software.php", 1, "", 1
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub mnuUploadOfflineAudits_Click()
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    
    Dim sameDb As Boolean
    Dim audits As Long, Files As Long, notes As Long
    Dim n As String
    Dim pid As Long
    
    On Error GoTo hell
      
    If cnDistro.State <> 0 Then cnDistro.Close
    cnDistro.Open
    
    If Err.Number <> 0 Then
        MsgBox "You seem to have lost your connection to distributed server"
        Exit Sub
    End If
    
    QueryDsn "Insert into tblauditlog(clsid) Values('sanitycheck')"
    Set rs = cn.Execute("Select * from tblauditlog where clsid='sanitycheck'")
    If Not rs.EOF Then sameDb = True
    QueryDsn "Delete from tblauditlog where clsid='sanitycheck'"
    
    If sameDb Then
        MsgBox "Your distributed server DSN is still pointing to your local database" & vbCrLf & "we can not export data to ourselves!", vbExclamation
        Exit Sub
    End If
    
    Set rs = cn.Execute("Select * from tblauditlog")
    If rs.EOF Then GoTo message
    
    While Not rs.EOF
        audits = audits + 1
        
        n = rs!notes
        If IsNull(n) Then n = ""
        
        DsnInsert "tblauditlog", "clsid,progid,auditor,sdate,version,crashs,tests,notes", _
                       rs!clsid, rs!ProgID, rs!auditor, rs!sdate, rs!version, _
                       rs!crashs, rs!tests, n

        rs.MoveNext
        
    Wend
    
    cn.Execute "Delete from tblauditlog"
    
    Set rs = cn.Execute("Select * from tblscripts")
    While Not rs.EOF
        Files = Files + 1
        
        n = rs!notes
        If IsNull(n) Then n = ""
        
        DsnInsert "tblscripts", "clsid,script,notes,auditor,sdate", _
                        rs!clsid, rs!script, n, rs!auditor, rs!sdate
        
        pid = QueryDsn("Select autoid from tblscripts where clsid='" & rs!clsid & "' and auditor='" & rs!auditor & "' order by autoid desc")!autoid
        If pid < 1 Then Exit Sub
        
        Set rs2 = cn.Execute("Select * from tblexceptions where pid=" & rs!autoid)
        While Not rs2.EOF
            DsnInsert "tblexceptions", "pid,sdata,disasm,address,exception", _
                         pid, rs2!sdata, rs2!Disasm, rs2!address, rs2!exception
            rs2.MoveNext
        Wend
                          
        rs.MoveNext
        
    Wend
        
    cn.Execute "Delete from tblscripts"
    cn.Execute "Delete from tblexceptions"
    
    Set rs = cn.Execute("Select * from tbldistnotes")
    While Not rs.EOF
        notes = notes + 1
        
        n = rs!notes
        If IsNull(n) Then n = ""
        
        DsnInsert "tbldistnotes", "auditor,sdate,notes,catagory", _
                    rs!auditor, rs!sdate, n, rs!catagory
                    
        rs.MoveNext
    Wend
    
    cn.Execute "Delete from tbldistnotes"
    
    Dim msg() As String
    push msg, "Export to master database successful: " & vbCrLf
    push msg, "Audits: " & audits
    push msg, "Files:  " & Files
    push msg, "Notes:  " & notes
    
    MsgBox Join(msg, vbCrLf), vbInformation
    
    
Exit Sub
message:
        MsgBox "You do not have any locally stored audits in your access database." & vbCrLf & _
                "" & vbCrLf & _
                "This feature is designed to allow you to save your work while on the " & vbCrLf & _
                "road by setting your COMRaider DSN to point to your local database" & vbCrLf & _
                "so you can use the distributed features on your own." & vbCrLf & _
                "" & vbCrLf & _
                "Once you get back to the lab, and reaim your dsn to be pointing to the" & vbCrLf & _
                "master server again, you can use this feature to upload the results to the " & vbCrLf & _
                "main distributed database" & vbCrLf

Exit Sub

hell:  MsgBox "Error in Export Data: " & Err.Description

End Sub

Private Sub mnuVCPLink_Click()
    On Error Resume Next
    ShellExecute 0, "Open", "http://labs.idefense.com/vcp.php", 1, "", 1
End Sub

Private Sub tmrKeepalive_Timer()
    ticks = ticks + 1
    If ticks > (60 * 5) Then '5 minutes
        On Error Resume Next
        ticks = 0
        cnDistro.Close
        Err.Clear
        cnDistro.Open
        If Err.Number <> 0 Then
            Options.DistributedMode = 0
            tmrKeepalive.Enabled = False
        End If
    End If
End Sub

Private Sub MDIForm_Load()

    With pictStartOff
        pictStartOn.Move .Left, .top
    End With
    
    On Error Resume Next
    ws(0).Close
    ws(0).LocalPort = 0 'let it choose a free one
    ws(0).Listen
    
    SaveSetting "COMRaider", "COMRaider", "port", ws(0).LocalPort
    
    Err.Clear
    Load ws(1)
    
    If Err.Number > 0 Then
        Me.caption = Me.caption & "  - Could not start local webserver on port " & ws(0).LocalPort
    End If
    
    Dim pth As String
    
    pth = App.path & IIf(IsIde, "\..", "") & "\comraider2.mdb"
    If Not fso.FileExists(pth) Then
       pth = App.path & IIf(IsIde, "\..", "") & "\comraider.mdb"
       If Not fso.FileExists(pth) Then
           MsgBox "Could not find " & pth
           End
       End If
    End If
     
     
    cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft " & _
                          "Access Driver (*.mdb)};DBQ=" & pth & ";"
    
    If cn.State = 0 Then cn.Open
    
    LoadOptions
        
End Sub

Private Sub pictStartOn_Click()
    On Error Resume Next
    Dim f As Form
    Dim x
    Dim ok() As String
    
    ok = Split("MDIForm1,frmAuditLogs,frmDistNotes", ",")
    
    For Each f In Forms
        For Each x In ok
            If x = f.Name Then GoTo nextOne
        Next
        Unload f
nextOne:
    Next
    frmLoadFile.Show
    
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrCheese_Timer()
    Dim p As POINTAPI
    GetCursorPos p
    If WindowFromPoint(p.x, p.y) <> pictStartOn.hwnd Then
        tmrCheese.Enabled = False
        pictStartOn.Visible = False
    End If
End Sub

Private Sub pictStartOff_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    pictStartOn.Visible = True
    tmrCheese.Enabled = True
End Sub
 
Private Sub MDIForm_Resize()
    On Error Resume Next
    Dim minW As Long
    Dim minH As Long
    Dim maxw As Long
 
    If Not setbg Then Picture1.BackColor = &H6D4623 '&H5A3008
    
    minW = Picture2.Width
    minH = 7725 + Picture1.Height + 300
    
    If Me.Width < minW Then Me.Width = minW
    If Me.Height < minH Then Me.Height = minH
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim f As Form
    On Error Resume Next
    SaveOptions
    For Each f In Forms
        Unload f
    Next
    End
End Sub



Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    ws(1).Close
    ws(1).Accept requestID
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim req As String
    Dim resp() As String
    Dim tmp() As String
    
    On Error Resume Next
    ws(Index).GetData req
    
    push resp, "HTTP/1.1 200 OK"
    push resp, "Server: COMRaider v" & App.Major & "." & App.Minor & "." & App.Revision
    push resp, "Connection: Close"
    push resp, "Transfer-Encoding: none"
    push resp, "Content-Type: text/html"
    push resp, "Content-Length: " & Len(webPage)
    push resp, ""
    push resp, webPage

    ws(Index).SendData Join(resp, vbCrLf)
    
End Sub


Sub ShowWebPageForGUID(sGuid As String, methName As String)
    Dim urlPath As String
    Dim tmp() As String
    
    On Error GoTo hell
    
    txtGUID = Replace(Replace(sGuid, "{", Empty), "}", Empty)
    
    push tmp, "<html>"
    push tmp, "<h2>This page has object guid: " & txtGUID & " embedded in it<br><br>"
    push tmp, "If its not safe for scripting then you will see an ActiveX warning<br><br>"
    push tmp, "This script will try to call a method on its interface at which time"
    push tmp, "IE will do the security checks. No arguments are given so dont worry about the error"
    push tmp, ""
    push tmp, "<object classid='clsid:" & txtGUID & "' id='target'></object>"
    push tmp, "<script language=vbs>"
    'push tmp, "Msgbox TypeName(target)"
    'push tmp, "On Error resume next"
    push tmp, "target." & methName
    push tmp, "</script>"
    webPage = Join(tmp, vbCrLf)
    
    urlPath = "http://127.0.0.1:" & ws(0).LocalPort & "/test.html"
    Shell """" & Options.IEPath & """ """ & urlPath & """", vbNormalFocus
        
    Exit Sub
hell:
        MsgBox Err.Description
End Sub

Sub TestCrashPage(wsfFilePath As String)
    
    Dim urlPath As String
    Dim tmp() As String
    
    On Error GoTo hell
    '<?XML version="1.0" standalone="yes" ?>
    '<package><job id="DoneInVBS"><?job debug="true"?>
    '<object id="target" classid="clsid:D7A7D7C3-D47F-11D0-89D3-00A0C90833E6"/>
    '<script language="VBScript">
    'WScript.Echo "This is VBScript"
    'WScript.Echo TypeName(Target)
    '</script></job></package>
    
    If Not fso.FileExists(wsfFilePath) Then Exit Sub
    
    webPage = fso.ReadFile(wsfFilePath)
    webPage = Replace(webPage, "{", Empty)
    webPage = Replace(webPage, "}", Empty)
    webPage = Replace(webPage, "<?XML version='1.0' standalone='yes' ?>", "<html>")
    webPage = Replace(webPage, "<package><job id='DoneInVBS' debug='false' error='true'>", "Test Exploit page")
    webPage = Replace(webPage, "</job></package>", Empty)
    webPage = Replace(webPage, "/>", "></object>")
    
    urlPath = "http://127.0.0.1:" & ws(0).LocalPort & "/test.html"
    Shell """" & Options.IEPath & """ """ & urlPath & """", vbNormalFocus
        
    Exit Sub
hell: MsgBox Err.Description
End Sub

