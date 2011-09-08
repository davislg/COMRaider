VERSION 5.00
Begin VB.Form frmAdvScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Object Scan"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopyList 
      Caption         =   "Copy List"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   2880
      Width           =   1035
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Scanning"
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5820
      Top             =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Hangs"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   5535
   End
End
Attribute VB_Name = "frmAdvScan"
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

Dim ticks As Long
Dim cnt  As Long
Dim abort As Boolean

Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Sub cmdAbort_Click()
    abort = True
    Command1.Enabled = True
End Sub

Private Sub cmdCopyList_Click()
    Dim tmp()
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To List1.ListCount
        push tmp, List1.List(i)
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(tmp, vbCrLf)
    MsgBox "List Copied", vbInformation
    
End Sub

Sub BuildForGuidList(tmp() As String, Optional always As Boolean = False)
    
    Dim v
    Dim i As Long
    Dim cnt As Long
    Dim processed As Long
    Dim startCount As Long
    Dim endCount As Long
    Dim pth As String
    
    On Error Resume Next
    
    Me.Visible = True
    Command1.Visible = False
    abort = False
    
    pth = """" & App.path & IIf(IsIde, "\..", "") & "\builddb.exe"""
    startCount = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
    
    For Each v In tmp
          
        If always Then
            cnt = 1
        Else
            cnt = cn.Execute("Select count(autoid) as cnt from tblScanned where clsid='" & v & "'")!cnt
        End If
        
        If cnt = 0 Then
            Label1 = "Testing Clsid: " & v
            Label1.Refresh
            cn.Execute "Insert into tblscanned(clsid) values('" & v & "')"
            ShellnWait pth, CStr(v)
            processed = processed + 1
        End If
        
        cnt = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
        
        DoEvents
        Sleep 30
        
        Me.caption = i & "/" & UBound(tmp) & "  -  " & cnt & " Guids in table"
        Me.Refresh
        i = i + 1
        
        If abort Then Exit For

    Next
    
    endCount = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
    
    On Error Resume Next
    
    MsgBox "Total New Clsid's found: " & processed & vbCrLf & _
           "Num Added Implementing Safety Options: " & (endCount - startCount), vbInformation
           
    
End Sub

Private Sub Command1_Click()
    
    Dim tmp() As String
    Dim v
    Dim i As Long
    Dim cnt As Long
    Dim processed As Long
    Dim startCount As Long
    Dim endCount As Long
    
    Dim pth As String
    
    abort = False
    reg.hive = HKEY_CLASSES_ROOT
    tmp = reg.EnumKeys("\CLSID")
    Command1.Enabled = False
    
    pth = """" & App.path & IIf(IsIde, "\..", "") & "\builddb.exe"""
    startCount = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
    
    For Each v In tmp
          
        cnt = cn.Execute("Select count(autoid) as cnt from tblScanned where clsid='" & v & "'")!cnt
        
        If cnt = 0 Then
            Label1 = "Testing Clsid: " & v
            Label1.Refresh
            cn.Execute "Insert into tblscanned(clsid) values('" & v & "')"
            ShellnWait pth, CStr(v)
            processed = processed + 1
        End If
        
        cnt = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
        
        DoEvents
        Sleep 30
        
        Me.caption = i & "/" & UBound(tmp) & "  -  " & cnt & " Guids in table"
        Me.Refresh
        i = i + 1
        
        If abort Then Exit For

    Next
    
    endCount = cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt
    
    On Error Resume Next
    
    MsgBox "Total New Clsid's found: " & processed & vbCrLf & _
           "Num Added Implementing Safety Options: " & (endCount - startCount), vbInformation
    
        
End Sub
 


 

Private Sub Form_Load()
    Me.Icon = MDIForm1.Icon
End Sub

Private Sub Timer1_Timer()
    
    ticks = ticks + 1
    If ticks > 10 Then
        ticks = 0
        Timer1.Enabled = False
    End If
    
End Sub

Sub ShellnWait(cmdLine As String, clsid As String)
    Dim status As Long, h As Long, pid As Long
    
    On Error GoTo hell
    pid = Shell(cmdLine & " " & clsid)
        
    ticks = 0
    Timer1.Enabled = True
    h = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    GetExitCodeProcess h, status
    
    Do While status = STILL_ACTIVE
        DoEvents
        Sleep 60
        GetExitCodeProcess h, status
        If Not Timer1.Enabled Then
            TerminateProcess h, 0
            List1.AddItem clsid
            Exit Do
        End If
    Loop
    
    CloseHandle h

hell:

End Sub
    

Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

