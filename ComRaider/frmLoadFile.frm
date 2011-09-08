VERSION 5.00
Begin VB.Form frmLoadFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select COM Server"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6750
   Begin VB.OptionButton optProgIDScan 
      Caption         =   "Search by ProgID"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1320
      Width           =   2355
   End
   Begin VB.OptionButton Option5 
      Caption         =   "View controls with KillBit set "
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1620
      Width           =   2415
   End
   Begin VB.OptionButton optDownloadShared 
      Caption         =   "View shared fuzz files from distributed audits"
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   2580
      Width           =   4035
   End
   Begin VB.OptionButton optScanDirectory 
      Caption         =   "Scan a directory for registered COM servers"
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   660
      Width           =   4035
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Select previously generated fuzz file to test"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2220
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   120
      Picture         =   "frmLoadFile.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton cndNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5340
      TabIndex        =   4
      Top             =   3060
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Choose from controls that should be loadable in IE"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Manually Enter the GUID "
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1020
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Choose ActiveX dll or ocx file directly"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1 - Select the COM Server you wish to test. "
      Height          =   255
      Left            =   1980
      TabIndex        =   0
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "frmLoadFile"
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


Private Sub cndNext_Click()
    Dim f As String, ff As String
    Dim Files() As String
    Dim key As String
    
    If Option1.value Then 'select file
        f = dlg.OpenDialog(AllFiles)
        If Len(f) = 0 Then Exit Sub
        frmtlbViewer.LoadFile f
        Unload Me
        
    ElseIf Option2.value Then 'load from guid
        
        ff = Trim(InputBox("Enter GUID you wish to analyze")) ', , "{4CECCEB2-8359-11D0-A34E-00AA00BDCDFD}")) '"05589FA1-C356-11CE-BF01-00AA0055595A"))
        If Len(ff) = 0 Then Exit Sub
        
        reg.hive = HKEY_CLASSES_ROOT
        
        If Right(ff, 1) <> "}" Then ff = ff & "}"
        If Left(ff, 1) <> "{" Then ff = "{" & ff
        
        f = "\CLSID\" & ff
        
        If reg.keyExists(f) Then
            f = f & "\InProcServer32"
            If reg.keyExists(f) Then
                f = reg.ReadValue(f, "")
                f = StripQuotes(f)
                If fso.FileExists(f) Then frmtlbViewer.LoadFile f, ff
            Else
                MsgBox "Could not find its InProcServer32 entry", vbInformation
            End If
        Else
            MsgBox "Could not locate this GUID on your system", vbInformation
        End If
        
    ElseIf Option4.value Then  'load test files and goto fuzz ui
        Dim tmp() As String
        Dim X
        
        f = LCase(dlg.FolderDialog("C:\COMRaid", Me.hwnd))
        If Len(f) = 0 Then Exit Sub
        
        If f = "c:\" Or InStr(f, LCase(Environ("WINDIR"))) > 0 Then
            MsgBox "This feature actually recursivly loads all wsf files in an below the directory you specify. Therefore you really dont want to load these directories", vbInformation
            Exit Sub
        End If
            
        tmp() = fso.GetFolderFiles(f, "wsf", True, True)
        
        If AryIsEmpty(tmp) Then
            MsgBox "Sorry no wsf files found", vbInformation
            Exit Sub
        End If
        
        'damn filter :(
        For Each X In tmp
            If InStr(X, ".wsf") > 0 Then push Files, X
        Next
        
        If AryIsEmpty(Files) Then
            MsgBox "Could not locate any wsf fuzz scripts in this directory!", vbInformation
            Exit Sub
        End If
        
       frmCrashMon.LoadFileList Files, True
       Unload Me
        
    ElseIf optScanDirectory.value Then
        
        f = dlg.FolderDialog
        If Len(f) = 0 Then Exit Sub
        If Not frmScanDir.ShowServersForPath(f) Then
            MsgBox "No COM Servers found in " & f, vbInformation
        Else
            Unload Me
        End If
        
    ElseIf optDownloadShared.value Then
        
        frmDownloadShared.Visible = True
        Unload Me
    
    ElseIf Option5.value Then
        
        frmKillBits.DisplayKillBitList
        Unload Me
        
    ElseIf optProgIDScan.value Then
        
        frmProgIDLookup.Visible = True
        Unload Me
        
    Else
        'select from safe for scripting list
        frmSafeForScripting.Visible = True
        Unload Me
        
    End If
    
End Sub

Sub Form_Load()
    Me.Icon = MDIForm1.Icon
    optDownloadShared.Enabled = Options.DistributedMode
End Sub

Private Sub lblOptions_Click()
    frmOptions.Show
End Sub
