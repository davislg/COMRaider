VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure Options"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Fuzzing Options "
      Height          =   3915
      Left            =   1740
      TabIndex        =   3
      Top             =   60
      Width           =   6855
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   6180
         TabIndex        =   28
         Top             =   3540
         Width           =   615
      End
      Begin VB.TextBox txtSymbolDir 
         Height          =   285
         Left            =   2640
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         Top             =   3540
         Width           =   3435
      End
      Begin VB.CheckBox chkLoadSymbols 
         Caption         =   "Use Symbols       Symbol Path"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3540
         Width           =   2475
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   255
         Left            =   6180
         TabIndex        =   25
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox txtApiTriggers 
         Height          =   315
         Left            =   2640
         TabIndex        =   24
         Top             =   3120
         Width           =   3435
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   6180
         TabIndex        =   22
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox txtBrowser 
         Height          =   285
         Left            =   1020
         TabIndex        =   21
         Top             =   1020
         Width           =   4935
      End
      Begin VB.TextBox txtApiFilter 
         Height          =   285
         Left            =   2640
         TabIndex        =   19
         Top             =   2820
         Width           =   4095
      End
      Begin VB.CheckBox chkUseApiLogger 
         Caption         =   "Use Api Logger"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2820
         Width           =   1515
      End
      Begin VB.CheckBox chkSimpleRegScanner 
         Caption         =   "Use Simple Registry Scanning mode to find safe objects (fast)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   4635
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkUseDistributedAuditing 
         Caption         =   "Use Distributed Auditng "
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1995
      End
      Begin VB.CheckBox chkOnlyDefaultInterface 
         Caption         =   "Only show the default Interface for classes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtDebugger 
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6180
         TabIndex        =   5
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox chkAllowObjRefs 
         Caption         =   "Allow fuzzing of functions that take Object type arguments"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   4395
      End
      Begin VB.Label Label8 
         Caption         =   "ApiTriggers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Browser"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Log Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3480
         MouseIcon       =   "frmOptions.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2160
         Width           =   195
      End
      Begin VB.Label Label4 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4860
         MouseIcon       =   "frmOptions.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4620
         MouseIcon       =   "frmOptions.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1800
         Width           =   195
      End
      Begin VB.Label Label3 
         Caption         =   "UserName"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblTestDSN 
         Caption         =   "Test Dsn Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2220
         MouseIcon       =   "frmOptions.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Debugger"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1635
      Left            =   60
      Picture         =   "frmOptions.frx":0C28
      ScaleHeight     =   1575
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4020
      Width           =   1035
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   7500
      TabIndex        =   0
      Top             =   4020
      Width           =   1035
   End
End
Attribute VB_Name = "frmOptions"
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
Dim tmp As String


Private Sub cmdBrowse_Click(Index As Integer)
    Dim f As String
    
    If Index < 2 Then
        f = dlg.OpenDialog(exeFiles, , "Choose your debugger")
    Else
         f = dlg.FolderDialog()
    End If
    
    If Len(f) = 0 Then Exit Sub
    f = Replace(f, Chr(0), "")
    
    If Index = 0 Then
        txtDebugger = f
    ElseIf Index = 1 Then
        txtBrowser = f
    Else
        txtSymbolDir = f
    End If
        
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
        
    txtUser = Replace(txtUser, "'", Empty)
    
    With Options
        .ExternalDebugger = txtDebugger
        .IEPath = txtBrowser
        .DistributedMode = IIf(chkUseDistributedAuditing.value = 1, True, False)
        .UserName = txtUser
        .AllowObjArgs = IIf(chkAllowObjRefs.value = 1, True, False)
        .OnlyDefInf = IIf(chkOnlyDefaultInterface.value = 1, True, False)
        .UseSimpleRegScanner = IIf(chkSimpleRegScanner.value = 1, True, False)
        .UseApiLogger = IIf(chkUseApiLogger.value = 1, True, False)
        .ApiFilters = txtApiFilter
        .ApiTriggers = txtApiTriggers
        .UseSymbols = IIf(chkLoadSymbols.value = 1, True, False)
        .SymPath = txtSymbolDir
    End With
    
    SaveOptions
    Unload Me
    
End Sub



Private Sub cmdTest_Click()
    tmp = InputBox("Enter string to test against current triggers", , tmp)
    If Len(tmp) = 0 Then Exit Sub
    If LikeAnyOfThese(tmp, txtApiTriggers) Then
        MsgBox "Match was successful", vbInformation
    Else
        MsgBox "Did not Match", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = MDIForm1.Icon
    
    With Options
        txtDebugger = Replace(.ExternalDebugger, Chr(0), "")
        chkUseDistributedAuditing.value = IIf(.DistributedMode, 1, 0)
        chkUseDistributedAuditing.Enabled = MDIForm1.tmrKeepalive.Enabled
        chkAllowObjRefs = IIf(.AllowObjArgs, 1, 0)
        chkOnlyDefaultInterface.value = IIf(.OnlyDefInf, 1, 0)
        chkSimpleRegScanner.value = IIf(.UseSimpleRegScanner, 1, 0)
        chkUseApiLogger.value = IIf(.UseApiLogger, 1, 0)
        txtUser = .UserName
        txtApiFilter = .ApiFilters
        txtBrowser = .IEPath
        txtApiTriggers = .ApiTriggers
        chkLoadSymbols.value = IIf(.UseSymbols, 1, 0)
        txtSymbolDir = .SymPath
        
    End With
    
    
    
End Sub

Private Sub Label2_Click()

    MsgBox "If a method has any non-default arguments of type 'Object' it will be" & vbCrLf & _
            "marked as a non-fuzzable function." & vbCrLf & _
            "" & vbCrLf & _
            "While Comraider does not support fuzzing object variables, if you check this" & vbCrLf & _
            "option it can try to run the function with a default object variable of 'Nothing'" & vbCrLf & _
            "Not ideal, but it may open some doors if processing was done before object" & vbCrLf & _
            "check. (Working on way to try to dynamically create live objects for args see src)", vbInformation
            
End Sub

Private Sub Label4_Click()

    MsgBox "This will be used for trying to find controls that could be loaded in IE." & vbCrLf & _
            "" & vbCrLf & _
            "If this is unchecked, then COMRaider will use an alternative method where it " & vbCrLf & _
            "tries to load each COM object registered on your machine and try to query it" & vbCrLf & _
            "for the IObjectSafety Interface to get more information." & vbCrLf & _
            "" & vbCrLf & _
            "The reason this more thorough test is not default is because it can take" & vbCrLf & _
            "quite a long time to complete." & vbCrLf & _
            "" & vbCrLf & _
            "If you uncheck this option make sure you understand this.", vbInformation
            
End Sub

Private Sub Label5_Click()
    MsgBox "Scripting clients such as Wscript which runs the wsf fuzz files can only" & vbCrLf & _
            "call functions on a COM Objects default interface." & vbCrLf & _
            "" & vbCrLf & _
            "Since COMRaider is designed to mainly find exploitable conditions in" & vbCrLf & _
            "COM Objects which can be launched from a scriptable web browser" & vbCrLf & _
            "this fine for our testing model. " & vbCrLf & _
            "" & vbCrLf & _
            "This option allows you to at least view functions that may reside on non-" & vbCrLf & _
            "default interfaces", vbInformation
End Sub

Private Sub Label6_Click()
    MsgBox "Comma delimited list of string fragments that if found = ignore" & vbCrLf & vbCrLf & _
            "Leave this field blank if you do not want to use it", vbInformation
End Sub

Private Sub Label8_Click()
     MsgBox "Comma delimited list of strings that if found = trigger alert (Matchs support * wildcards)" & vbCrLf & vbCrLf & _
            "Leave this field blank if you do not want to use it" & vbCrLf & _
            "Double click textbox to edit in larger window", vbInformation
End Sub

Private Sub lblTestDSN_Click()
    On Error Resume Next
    cnDistro.Close
    
    Err.Clear
    cnDistro.Open
    
    If Err.Number <> 0 Then 'data source not found or could not connect
        chkUseDistributedAuditing.value = 0
        Options.DistributedMode = False
    Else
        Options.DistributedMode = True
    End If
    
    With MDIForm1
        chkUseDistributedAuditing.Enabled = Options.DistributedMode
        .tmrKeepalive.Enabled = Options.DistributedMode
        .mnuAuditLogs.Enabled = Options.DistributedMode
        .mnuDistNotes.Enabled = Options.DistributedMode
        .mnuUploadOfflineAudits.Enabled = Options.DistributedMode
    End With
                        
    MsgBox IIf(Options.DistributedMode, "DSN Connection was successful", "DSN Connection Failed"), vbInformation

End Sub

Private Sub txtApiTriggers_DblClick()
    Dim x As String
    x = frmMsg.GetData(, txtApiTriggers)
    txtApiTriggers = Replace(x, vbCrLf, "")
End Sub

Private Sub txtDebugger_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtDebugger = Data.Files(1)
End Sub

Private Sub txtSymbolDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If fso.FolderExists(Data.Files(1)) Then
        txtSymbolDir = Data.Files(1)
    Else
        MsgBox "Only drop folders in here", vbInformation
    End If
End Sub
