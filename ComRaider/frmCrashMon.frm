VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCrashMon 
   Caption         =   "Debugger Interface"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9345
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   4335
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   4395
      Begin VB.Frame Frame2 
         Height          =   2115
         Left            =   0
         TabIndex        =   16
         Top             =   -60
         Width           =   4275
         Begin VB.TextBox txtAuditNotes 
            Height          =   285
            Left            =   1020
            TabIndex        =   22
            Text            =   "nothing of interest"
            Top             =   1380
            Width           =   1515
         End
         Begin VB.CheckBox chkUploadAudits 
            Caption         =   "Upload Audit Log"
            Height          =   195
            Left            =   2580
            TabIndex        =   20
            Top             =   1440
            Width           =   1635
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            Height          =   315
            Left            =   3000
            TabIndex        =   17
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmCrashMon.frx":0000
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label3 
            Caption         =   "Audit Notes"
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   1155
            Left            =   1020
            TabIndex        =   18
            Top             =   240
            Width           =   3195
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5115
      Left            =   0
      TabIndex        =   1
      Top             =   2100
      Width           =   9255
      Begin MSComctlLib.ListView lvApiLog 
         Height          =   975
         Left            =   0
         TabIndex        =   19
         Top             =   2100
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Api Log"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   7500
         TabIndex        =   6
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Begin Fuzzing"
         Height          =   375
         Left            =   7500
         TabIndex        =   12
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Timer tmrTimeout 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   1860
         Top             =   4380
      End
      Begin VB.Timer tmrWindowKiller 
         Enabled         =   0   'False
         Interval        =   60
         Left            =   1860
         Top             =   4800
      End
      Begin VB.CheckBox chkWindowKiller 
         Caption         =   "Close Popups"
         Height          =   195
         Left            =   4140
         TabIndex        =   10
         Top             =   4440
         Width           =   1275
      End
      Begin VB.CheckBox chkTimeout 
         Caption         =   "Kill hung Processes"
         Height          =   255
         Left            =   2340
         TabIndex        =   9
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CheckBox chkPause 
         Caption         =   "Pause"
         Height          =   195
         Left            =   6060
         TabIndex        =   8
         Top             =   4440
         Width           =   795
      End
      Begin VB.CheckBox chkDeleteDuds 
         Caption         =   "Delete duds"
         Height          =   255
         Left            =   2340
         TabIndex        =   5
         Top             =   4740
         Width           =   1215
      End
      Begin VB.OptionButton optDebugContinue 
         Caption         =   "dbg_continue"
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   4860
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optDebugNotHandled 
         Caption         =   "dbg_not_handled"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   4620
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkSaveonlyExceptions 
         Caption         =   "Save only exceptions"
         Height          =   255
         Left            =   4140
         TabIndex        =   2
         Top             =   4740
         Width           =   1875
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   4080
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvMsg 
         Height          =   1215
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Exception"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Module"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Instruction"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvWindow 
         Height          =   915
         Left            =   0
         TabIndex        =   13
         Top             =   1200
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1614
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Class"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Caption"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvDbg 
         Height          =   915
         Left            =   0
         TabIndex        =   23
         Top             =   3120
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   1614
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Debug Strings"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "On Exception use"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4380
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Result"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Exceptions"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Windows"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ApiHits"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuDebugTest 
         Caption         =   "Edit for debug test"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFile 
         Caption         =   "View File"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save To"
      End
      Begin VB.Menu mnuCopyFileName 
         Caption         =   "Copy File Name"
      End
      Begin VB.Menu mnuSpacerxx 
         Caption         =   "-"
      End
      Begin VB.Menu cmdShowTlb 
         Caption         =   "View Tlb"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTestInIE 
         Caption         =   "Test Exploit in IE"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUploadAudits 
         Caption         =   "Upload Audit Results"
      End
      Begin VB.Menu mnuUploadFileToDB 
         Caption         =   "Upload File to Database"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaunchNormal 
         Caption         =   "Launch Normal"
      End
      Begin VB.Menu mnuLaunchExternalDebugger 
         Caption         =   "Launch in Olly"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteSelected 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuSelNoWindows 
         Caption         =   "Select No Windows"
      End
      Begin VB.Menu mnuSelectNoCrash 
         Caption         =   "Select No Exceptions"
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchApiLogs 
         Caption         =   "Search ApiLogs"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCrashDetails 
         Caption         =   "View Details"
      End
   End
End
Attribute VB_Name = "frmCrashMon"
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

Public WithEvents Crash As CCrashMon
Attribute Crash.VB_VarHelpID = -1

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim AuditedClsIDs As New Collection 'of CAuditEntry
Private curLogger As CLogger
Dim liWindow As ListItem
Dim liFile As ListItem
Dim liMsg As ListItem
Dim abort As Boolean
Dim auditsUploaded As Boolean

Dim Exceptions As Long
Dim Windows As Long
Dim ApiHits As Long


Private Sub chkDeleteDuds_Click()
    chkSaveonlyExceptions.Enabled = CBool(chkDeleteDuds.value)
End Sub

Private Sub chkPause_Click()
    tmrWindowKiller.Enabled = IIf(chkPause.value = 1, True, False)
End Sub

Private Sub cmdOk_Click()
    
    If chkUploadAudits.Visible And chkUploadAudits.value = 1 And Options.DistributedMode Then
        DoAuditUpload txtAuditNotes.Text
    End If
        
    Picture1.Visible = False
    
End Sub

Private Sub cmdShowTlb_Click()
    
    If liFile Is Nothing Then
        MsgBox "Because this list can be loaded to fuzz multiple libraries at once, please select a file from the library you wish to view the tlb file for.", vbInformation
        Exit Sub
    End If
    
    Dim tmp() As String
    Dim s As String
    Dim clsid As String

    s = fso.ReadFile(liFile.Text)
    
    x = InStr(s, "classid='clsid:") + 15
    y = InStr(x, s, "'")
    clsid = Mid(s, x, y - x)
    clsid = "{" & clsid & "}"
    
    reg.hive = HKEY_CLASSES_ROOT
    
    s = "CLSID\" & clsid
    If Not reg.keyExists(s) Then GoTo generalFail
    
    s = s & "\InProcServer32"
    If Not reg.keyExists(s) Then GoTo generalFail
        
    s = reg.ReadValue(s, "")
    If fso.FileExists(s) Then
        frmtlbViewer.LoadFile s, clsid
    Else
        MsgBox "Could not find file: " & s
    End If
        
Exit Sub
generalFail:
    
    MsgBox "Could not find registry key for HKCR\" & s, vbInformation
    Exit Sub
    
End Sub

Private Sub cmdStop_Click()
    abort = True
    Crash.StopDbg
End Sub

Private Sub Crash_DebugString(msg As String)
    
    'If Len(Options.DbgTriggers) > 0 And LikeAnyOfThese(msg, Options.DbgTiggers) Then
        curLogger.DbgStrings.Add msg
        curLogger.DbgHits = curLogger.DbgHits + 1
    'End If
    
End Sub

Private Sub mnuSearchApiLogs_Click()
    
    Dim li As ListItem
    Dim results() As String
    Dim lg As CLogger
    Dim match As String
    Dim tmp As String
    
    On Error Resume Next
    
    match = InputBox("Enter comma delimited string of substrings to search for (supports * wildcards) ")
    If Len(match) = 0 Then Exit Sub
    match = "*" & match & "*"
    
    For Each li In lv.ListItems
        If IsObject(li.Tag) Then
            tmp = Empty
            Set lg = li.Tag
            tmp = lg.GetApiLogAsString
            If LikeAnyOfThese(tmp, match) Then
                push results, "File: " & li.Text & vbCrLf & tmp & vbCrLf & String(75, "-") & vbCrLf
            End If
        End If
    Next
            
    If AryIsEmpty(results) Then
        MsgBox "Sorry no matchs found", vbInformation
    Else
        frmMsg.Display Join(results, vbCrLf)
    End If
    
End Sub

Private Sub mnuUploadAudits_Click()
    DoAuditUpload
End Sub

Sub DoAuditUpload(Optional note As String = "-1")

    If auditsUploaded Then
        MsgBox "You have already uploaded this audit list", vbInformation
        Exit Sub
    End If
    
    If AuditedClsIDs.Count = 0 Then
        MsgBox "You have not yet audited any Clsids", vbInformation
        Exit Sub
    End If
    
    Dim e As CAuditEntry
    Dim ProgID As String
    Dim file As String
    Dim version As String
    Dim key1 As String, key2 As String
    Dim crashs As Long
    Dim default As String
        
    If note = "-1" Then
        default = IIf(Exceptions > 0, "", "nothing of interest")
        note = InputBox("Enter a note (if any) you would like applied to these audits in the log", , default)
    End If
    
    note = Replace(note, "'", Empty)
    If Len(note) = 0 Then note = " " 'keep from possible null data error

    
    Dim l As CLogger
    Dim li As ListItem
    
    'now we have run our audits so we can tally crashs
    'and add them to our cauditentry classes
    For Each li In lv.ListItems
        If Len(li.SubItems(2)) > 0 Then
            If li.SubItems(2) > 0 Then
                For Each e In AuditedClsIDs
                    If e.IncludesFile(li.Text) Then
                        e.AddExceptions li.SubItems(2)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    reg.hive = HKEY_CLASSES_ROOT
    
    For Each e In AuditedClsIDs
        ProgID = Empty
        file = Empty
        version = Empty
        
        key1 = "\CLSID\" & e.GUID & "\Progid"
        key2 = "\CLSID\" & e.GUID & "\InProcServer32"
        
        If reg.keyExists(key1) Then ProgID = reg.ReadValue(key1, "")
        
        If reg.keyExists(key2) Then
            file = reg.ReadValue(key2, "")
            If fso.FileExists(file) Then version = FileVersion(file)
        End If
        
        DsnInsert "tblauditlog", "clsid,progid,auditor,sdate,version,crashs,tests,notes", _
                    e.GUID, ProgID, Options.UserName, Format(Now, "m.d.yy"), _
                    version, e.Exceptions, e.Files.Count, note
                    
    Next
    
    MsgBox AuditedClsIDs.Count & " ClsID Records uploaded to server", vbInformation
    
    Dim f As Form
    Dim fname As String
    
    'update the displays of other forms which may be up
    For Each f In Forms
        fname = TypeName(f)
        If fname = "frmtlbViewer" Or fname = "frmSafeForScripting" Or fname = "frmScanDir" Then
            For Each e In AuditedClsIDs
                f.MarkGUIDAsAudited e.GUID
            Next
        End If
    Next
    
    auditsUploaded = True
    Set AuditedClsIDs = New Collection
    Me.ZOrder 0
    
End Sub

Private Sub Form_Load()
                
    Me.top = 0
    Me.Left = 0
    Me.Icon = MDIForm1.Icon
    SizeLV lv, True
    SizeLV lvMsg
    SizeLV lvWindow
    SizeLV lvApiLog
     
    cmdStop.top = cmdBegin.top
    cmdStop.Visible = False
    mnuUploadAudits.Visible = Options.DistributedMode
    mnuUploadFileToDB.Visible = Options.DistributedMode
    
    chkTimeout.value = GetSetting("ComRaider", "Settings", "chkTimeout", 1)
    chkWindowKiller.value = GetSetting("ComRaider", "Settings", "chkWindowKiller", 1)
    chkDeleteDuds.value = GetSetting("ComRaider", "Settings", "chkDeleteDuds", 1)
    chkSaveonlyExceptions.value = GetSetting("ComRaider", "Settings", "chkSaveonlyExceptions", 1)
    chkUploadAudits.value = GetSetting("ComRaider", "Settings", "chkUploadAudits", 1)
    
    If chkDeleteDuds.value = 0 Then chkSaveonlyExceptions.Enabled = False
    
    SaveSetting "ComRaider", "Settings", "hwnd", Me.hwnd
    
    Set Crash = New CCrashMon
    Crash.Initilize Me.hwnd
        
End Sub


Sub LoadFileList(Files() As String, Optional showTlbButton As Boolean = False)
    
    If AryIsEmpty(Files) Then
        MsgBox "No fuzz files to load", vbInformation
        Unload Me
        Exit Sub
    End If
    
    Dim x
    lv.ListItems.Clear
    For Each x In Files
        lv.ListItems.Add , , x
    Next
    
    auditsUploaded = False
    cmdShowTlb.Visible = showTlbButton
    Me.caption = Me.caption & "  " & UBound(Files) & " files loaded"
    Me.Visible = True
    
End Sub


Private Sub cmdBegin_Click()

    On Error Resume Next
    
    abort = False
    cmdStop.Visible = True
    Exceptions = 0
    Windows = 0
    ApiHits = 0
    
    If lv.ListItems.Count = 0 Then
        MsgBox "There are no files to fuzz"
        Exit Sub
    End If
    
    Dim li As ListItem
    Dim logger As CLogger
    Dim i As Long
    Dim cmd As String
    
    chkPause.value = 0
    'pb.Visible = True
    pb.Max = lv.ListItems.Count
    pb.value = 0
    
    'blank out any old results if they re refuzzing
    For Each li In lv.ListItems
        li.SubItems(1) = Empty
        li.SubItems(2) = Empty
        li.SubItems(3) = Empty
        li.SubItems(4) = Empty
    Next
    
    For Each li In lv.ListItems
        If fso.FileExists(li.Text) Then
            
            If chkPause.value = 1 Then tmrWindowKiller.Enabled = False
                
            While chkPause.value
                Sleep 60
                DoEvents
                If abort Then Exit For
            Wend
            
            If abort Then Exit For
            
            tmrWindowKiller.Enabled = (chkWindowKiller.value = 1)
            tmrTimeout.Enabled = False 'reset interval
            tmrTimeout.Enabled = (chkTimeout.value = 1)
 
            Set logger = New CLogger
            Set curLogger = logger
            Set li.Tag = logger
            Set logger.li = li
            li.EnsureVisible 'scroll new item it into view
            
            If Options.DistributedMode Then BuildAuditList li.Text
            
            Me.caption = "On file " & pb.value & "/" & pb.Max & "     " & Exceptions & " Exceptions    " & Windows & " Windows Closed"
            
            If InStr(li.Text, ".exe") > 0 Then 'debugging test
                cmd = li.Text
            Else
                cmd = "wscript.exe """ & li.Text & """"
            End If
            
            If Not Crash.LaunchProcess(cmd) Then
                li.SubItems(1) = "Debugger err: " & Crash.GetErr
            Else
                While Crash.isDebugging
                    DoEvents
                    Sleep 60
                    If abort Then Exit For
                    If chkTimeout.value = 1 Then
                        If Not tmrTimeout.Enabled Then
                            li.SubItems(1) = "Timeout"
                            Crash.StopDbg
                        End If
                    End If
                Wend
                If Crash.CausedCrash Then li.SubItems(1) = "Caused Exception"
                li.SubItems(2) = logger.Exceptions.Count
                li.SubItems(3) = logger.Windows.Count
                li.SubItems(4) = logger.ApiHits
                Exceptions = Exceptions + logger.Exceptions.Count
                Windows = Windows + logger.Windows.Count
                ApiHits = ApiHits + logger.ApiHits
            End If
            
            lv_ItemClick li  'display all captured data for item
            
        End If
        pb.value = pb.value + 1
    Next
    
    Crash.StopDbg
    'pb.Visible = False
    pb.value = 0
    cmdStop.Visible = False
    
    chkPause.value = 0
    tmrTimeout.Enabled = False
    tmrWindowKiller.Enabled = False
    
    Dim deleteIt As Boolean
    
    If chkDeleteDuds.value = 1 Then
top:
        For i = 1 To lv.ListItems.Count
            Set li = lv.ListItems(i)
            If IsObject(li.Tag) And Len(li.SubItems(1)) = 0 Then
                Set logger = li.Tag
                If logger.Exceptions.Count = 0 Then
                    deleteIt = True
                    If chkSaveonlyExceptions.value = 0 Then
                        'do not delete if windows were displayed
                        If logger.Windows.Count > 0 Then deleteIt = False
                        If logger.ApiHits > 0 Then deleteIt = False
                    End If
                    If Len(li.SubItems(1)) > 0 Then
                        'some info msg was logged do not delete
                        deleteIt = False
                    End If
                    If deleteIt Then
                        fso.DeleteFile li.Text
                        lv.ListItems.Remove i
                        GoTo top
                    End If
                    
                End If
            End If
        Next
        
    End If
    
    Dim tmp()
    push tmp, Space(10) & "Fuzzing Complete." & vbCrLf
    push tmp, Windows & "   Windows"
    push tmp, Exceptions & "   Exceptions "
    push tmp, ApiHits & "   ApiTriggers "
    
    
    txtAuditNotes.Visible = (Options.DistributedMode And (lv.ListItems.Count < 2))
    chkUploadAudits.Visible = (Options.DistributedMode And (lv.ListItems.Count < 2))
    Label3.Visible = (Options.DistributedMode And (lv.ListItems.Count < 2))
 
    Label2.caption = Join(tmp, vbCrLf)
    Picture1.Visible = True
    
    'we dont want a modal msgbox cause it might be up for hours if you left it
    'to fuinish unattended..if that happened the mdiform1.tmrkeepalive would
    'be frozen out the whole time andyou would loose your connection to the
    'myswl server, which means whenyou tried to upload your long azz audit results
    'you would get a big error and be quite pissed (this i know)
    
    
End Sub

Private Sub Crash_Crash(e As CException)
    curLogger.Exceptions.Add e
End Sub

Private Sub Crash_ApiLogMsg(msg As String)
    
    'trigger matchs override filters
    If Len(Options.ApiTriggers) > 0 And LikeAnyOfThese(msg, Options.ApiTriggers) Then
        curLogger.ApiLogs.Add msg
        curLogger.ApiHits = curLogger.ApiHits + 1
    Else
        If Not AnyOfTheseInstr(msg, Options.ApiFilters) Then
            curLogger.ApiLogs.Add msg
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
 
    If Me.Width < 9465 Then
        Me.Width = 9465
    Else
        lv.Width = Me.Width - lv.Left - 150
        Frame1.Width = Me.Width
        lvMsg.Width = lv.Width
        lvWindow.Width = lv.Width
        lvApiLog.Width = lv.Width
        lvDbg.Width = lv.Width
        pb.Width = lv.Width
    End If
        
    If Me.Height < 7695 Then
        Me.Height = 7695
    Else
        Frame1.top = Me.Height - Frame1.Height - 450
        lv.Height = Frame1.top - lv.top - 100
    End If
    
    SizeLV lv
    SizeLV lvMsg
    SizeLV lvWindow
    SizeLV lvApiLog
    SizeLV lvDbg
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If cmdStop.Visible Then
        MsgBox "Stop the current tests first", vbInformation
        Cancel = 1
        Exit Sub
    End If
        
    Crash.StopDbg 'just in case (if still debugging then fatal crash)
    Set Crash = Nothing 'make sure the subclasses tear down properly
    
    Sleep 60
    DoEvents
    
    SaveSetting "ComRaider", "Settings", "chkTimeout", chkTimeout.value
    SaveSetting "ComRaider", "Settings", "chkWindowKiller", chkWindowKiller.value
    SaveSetting "ComRaider", "Settings", "chkDeleteDuds", chkDeleteDuds.value
    SaveSetting "ComRaider", "Settings", "chkSaveonlyExceptions", chkSaveonlyExceptions.value
    SaveSetting "ComRaider", "Settings", "chkUploadAudits", chkUploadAudits.value
    
    If AuditedClsIDs.Count > 0 Then
        If MsgBox("Audited clsids are ready for upload would you like to do it now?", vbInformation + vbYesNo) = vbYes Then
            mnuUploadAudits_Click
        End If
    End If
    
End Sub

Private Sub Label1_Click()

    MsgBox "DBG_NOT_HANDLED = pass exception to client as expected" & vbCrLf & _
            "DBG_CONTINUE = act as if exception didnt happen and continue" & vbCrLf & _
            "" & vbCrLf & _
            "You can use DBG_CONTINUE to simulate the case where you" & vbCrLf & _
            "modified your exploit payload so that it did not crash at the first " & vbCrLf & _
            "offset shown to try to get to latter offsets. Only applicable where you could." & vbCrLf & _
            "craft your input to bypass initial crashs"

End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim li As ListItem
    Dim logger As CLogger
    Dim e As CException
    Dim w As CWindow
    Dim c As Collection
    Dim v As Variant
    
    lvMsg.ListItems.Clear
    lvWindow.ListItems.Clear
    lvApiLog.ListItems.Clear
    lvDbg.ListItems.Clear
    
    Set liFile = Item
    
    If Not IsObject(Item.Tag) Then Exit Sub
    
    Set logger = Item.Tag
    
    For Each e In logger.Exceptions
        Set li = lvMsg.ListItems.Add(, , Hex(e.ExceptionAddress))
        Set li.Tag = e
        li.SubItems(1) = Crash.NameForDebugEvent(e.ExceptionCode)
        li.SubItems(2) = e.CrashInModule
        li.SubItems(3) = e.Disasm
    Next
    
    For Each c In logger.Windows
        For Each w In c
            If w.Class <> "Button" And Len(w.caption) > 0 Then
                Set li = lvWindow.ListItems.Add(, , w.Class)
                li.SubItems(1) = w.caption
            End If
        Next
    Next
    
    For Each v In logger.ApiLogs
        lvApiLog.ListItems.Add , , CStr(v)
    Next
    
    For Each v In logger.DbgStrings
        lvDbg.ListItems.Add , , v
    Next
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If liFile Is Nothing Then Exit Sub
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lvMsg_DblClick()
    If liMsg Is Nothing Then Exit Sub
    
    Dim e As CException
    Dim msg() As String
    
    Set e = liMsg.Tag
    
    push msg, "Exception Code: " & Crash.NameForDebugEvent(e.ExceptionCode)
    push msg, "Disasm: " & Hex(e.ExceptionAddress) & vbTab & e.Disasm & vbTab & "(" & e.CrashInModule & ")" & vbCrLf
    push msg, e.Enviroment
    
    frmMsg.Display Join(msg, vbCrLf)

End Sub

Private Sub lvMsg_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liMsg = Item
End Sub

Private Sub lvMsg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup2
End Sub

Private Sub lvWindow_DblClick()
    If liWindow Is Nothing Then Exit Sub
    
    If Len(liWindow.SubItems(1)) > 0 Then
        MsgBox liWindow.SubItems(1)
    End If
    
End Sub

Private Sub lvWindow_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liWindow = Item
End Sub

Private Sub mnuCopyFileName_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText liFile.Text
    If Err.Number = 0 Then MsgBox "File named saved to clipboard", vbInformation
End Sub

Private Sub mnuDebugTest_Click()
    If liFile Is Nothing Then Exit Sub
    Dim x As String
    x = dlg.OpenDialog(exeFiles, , "Choose exe to test debugger with")
    If Len(x) = 0 Then Exit Sub
    liFile.Text = x
End Sub

Private Sub mnuDeleteSelected_Click()
    Dim li As ListItem
    
    On Error Resume Next
    
    If cmdStop.Visible And chkPause.value = 0 Then
        MsgBox "Wait for test to conclude or pause it", vbInformation
        Exit Sub
    End If
    
    Dim cnt As Long
    For Each li In lv.ListItems
        If li.Selected Then cnt = cnt + 1
    Next
    
    If cnt = 0 Then Exit Sub
    If MsgBox("Are you sure you want to delete these " & cnt & " files?", vbYesNo + vbInformation) = vbNo Then Exit Sub
    
    For cnt = lv.ListItems.Count To 1 Step -1
        If lv.ListItems(cnt).Selected Then
            If fso.FileExists(lv.ListItems(cnt).Text) Then
                fso.DeleteFile lv.ListItems(cnt).Text
            End If
            lv.ListItems.Remove cnt
        End If
    Next
    
End Sub

Private Sub mnuLaunchExternalDebugger_Click()
    
    If liFile Is Nothing Then Exit Sub
    
    If Len(Options.ExternalDebugger) = 0 Then
        If MsgBox("You have not set your external debugger , would you like to do this now?", vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
        frmOptions.Show 1, MDIForm1
        If Len(Options.ExternalDebugger) = 0 Then Exit Sub
        If Not fso.FileExists(Options.ExternalDebugger) Then Exit Sub
    End If
    
    Dim ws As String
    Dim cmd As String
    
    ws = Environ("Windir") & "\System32\wscript.exe"
    If Not fso.FileExists(ws) Then
        MsgBox "Could not locate wscript.exe? " & ws, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    cmd = """" & Options.ExternalDebugger & """ """ & ws & """ """ & liFile.Text & """"
    Shell cmd, vbNormalFocus
        
    If Err.Number > 0 Then
        MsgBox "Error launching cmd: " & vbCrLf & vbCrLf & cmd, vbInformation
    End If
        
End Sub

Private Sub mnuLaunchNormal_Click()
    On Error Resume Next
    Dim ws As String
    Dim cmd As String
    
    If liFile Is Nothing Then Exit Sub
    
    ws = Environ("Windir") & "\System32\wscript.exe"
    If Not fso.FileExists(ws) Then
        MsgBox "Could not locate wscript.exe? " & ws, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    cmd = """" & ws & """ """ & liFile.Text & """"
    Shell cmd, vbNormalFocus
        
    If Err.Number > 0 Then
        MsgBox "Error launching cmd: " & vbCrLf & vbCrLf & cmd, vbInformation
    End If

End Sub

Private Sub mnuSaveAs_Click()
    Dim f As String
    Dim d As String
    Dim e As CException
    Dim l As CLogger
    Dim tmp() As String
    
    On Error GoTo hell
    
    If liFile Is Nothing Then Exit Sub
    
    If Not fso.FileExists(liFile.Text) Then
        MsgBox "Could not locate file " & liFile.Text, vbInformation
        Exit Sub
    End If

    f = dlg.FolderDialog
    If Len(f) = 0 Then Exit Sub
    
    fso.Copy liFile.Text, f
    
    If IsObject(liFile.Tag) Then
        d = f & "\" & Replace(fso.FileNameFromPath(liFile.Text), ".wsf", ".txt")
        
        Set l = liFile.Tag
        
        For Each e In l.Exceptions
            push msg, "Exception Code: " & Crash.NameForDebugEvent(e.ExceptionCode)
            push msg, "Disasm: " & Hex(e.ExceptionAddress) & vbTab & e.Disasm & vbCrLf
            push msg, e.Enviroment
            push msg, ""
        Next
        
        If l.ApiLogs.Count > 0 Then
            push msg, "ApiLog" & vbCrLf & String(50, "-") & vbCrLf
            push msg, l.GetApiLogAsString
        End If
        
        If l.DbgStrings.Count > 0 Then
            push msg, "Debug String Log" & vbCrLf & String(50, "-") & vbCrLf
            push msg, l.GetDbgStrsAsString
        End If
        
        fso.WriteFile d, Join(msg, vbCrLf)
    End If
    
    MsgBox "Copy Complete", vbInformation
    
    Exit Sub
hell: MsgBox "Error: " & Err.Description
End Sub

Private Sub mnuSelectNoCrash_Click()
    Dim li As ListItem
    On Error Resume Next
    For Each li In lv.ListItems
        If li.SubItems(2) = 0 Then li.Selected = True
    Next
End Sub

Private Sub mnuSelNoWindows_Click()
    
    Dim li As ListItem
    On Error Resume Next
    For Each li In lv.ListItems
        If li.SubItems(3) = 0 Then li.Selected = True
    Next
    
End Sub

Private Sub mnuTestInIE_Click()
    If liFile Is Nothing Then Exit Sub
    MDIForm1.TestCrashPage liFile.Text
End Sub

Private Sub mnuUploadFileToDB_Click()
    
   Dim cnt As Integer
   Dim notes As String
   Dim li As ListItem
   Dim pid As Long
   Dim clsid As String
   Dim l As CLogger
   Dim e As CException
   Dim apilog As String
   Dim ProgID As String
   
   notes = InputBox("Enter your notes for these selected items (if any)", , "exploitable")
   
   reg.hive = HKEY_CLASSES_ROOT
   
   For Each li In lv.ListItems
        If li.Selected Then
            If fso.FileExists(li.Text) Then
                cnt = cnt + 1
                clsid = ExtractGUID(li.Text)
                ProgID = reg.ReadValue("\CLSID\" & clsid & "\Progid", "")
                
                apilog = Empty
                Set l = Nothing
                
                If IsObject(li.Tag) Then
                    Set l = li.Tag
                    apilog = l.GetApiLogAsString
                End If
                
                DsnInsert "tblscripts", "clsid,script,notes,auditor,sdate,apilog,progid", _
                            clsid, fso.ReadFile(li.Text), notes, Options.UserName, _
                            Format(Now, "m.d.yy"), apilog, ProgID
                            
                If IsObject(li.Tag) Then
                    If l.Exceptions.Count > 0 Then
                        pid = QueryDsn("Select autoid from tblscripts where clsid='" & clsid & "' and auditor='" & Options.UserName & "' order by autoid desc")!autoid
                        If pid > 0 Then
                            For Each e In l.Exceptions
                                DsnInsert "tblexceptions", "pid,sdata,disasm,address,exception", _
                                            pid, e.Enviroment, e.Disasm, e.ExceptionAddress, _
                                            Crash.NameForDebugEvent(e.ExceptionCode)
                            Next
                        End If
                    End If
                End If
            
            End If
        End If
   Next
    
   MsgBox cnt & " files uploaded to shared database", vbInformation
    
End Sub

Private Sub mnuViewCrashDetails_Click()
    lvMsg_DblClick
End Sub

Private Sub mnuViewFile_Click()
    On Error Resume Next
    Shell "notepad.exe """ & liFile.Text & """", vbNormalFocus
End Sub

Private Sub tmrTimeout_Timer()
    tmrTimeout.Enabled = False
End Sub

Private Sub tmrWindowKiller_Timer()
    
    Dim hwnd As Long
    Dim p As String
    Dim a As Collection
    Dim c As Collection
    Dim w1 As CWindow
    Dim w As CWindow
    Dim Class As String
 
'of course this one just doesnt work heh
'    Set a = EnumChildren(0)
'    For Each w1 In a
'        p = ProcessFromHwnd(w1.hwnd)
'        If InStr(1, p, "wscript.exe", vbTextCompare) > 0 Then
'            If w1.Class = "#32770" Then
'                Set c = EnumChildren(hwnd)
'                For Each w In c
'                    If w.Class = "Button" Then
'                        If LCase(w.caption) = "ok" Then
'                            curLogger.Windows.Add c
'                            PostMessage w.hwnd, WM_CLICK, 1, 1
'                            Exit Sub
'                        End If
'                    End If
'                Next
'            End If
'            Exit For
'        End If
'    Next
    
    
'below version does not work with screen saver active
    hwnd = GetForegroundWindow()
    p = ProcessFromHwnd(hwnd)

    If InStr(1, p, "wscript.exe", vbTextCompare) < 1 Then Exit Sub

    Class = GetClass(hwnd)
    Select Case Class

        Case "#32770" 'modal messagebox or windows dialog
                Set c = EnumChildren(hwnd)
                For Each w In c
                    If w.Class = "Button" Then
                        If LCase(w.caption) = "ok" Then
                            curLogger.Windows.Add c
                            PostMessage w.hwnd, WM_CLICK, 1, 1
                            Exit Sub
                        End If
                    End If
                Next

    End Select
        
 
 
    
End Sub

Sub BuildAuditList(WSFFile As String)
    Dim GUID As String
    Dim e As CAuditEntry
    
    GUID = ExtractGUID(WSFFile)
    If Len(GUID) = 0 Then Exit Sub
        
    If Not KeyExistsInCollection(AuditedClsIDs, GUID) Then
        Set e = New CAuditEntry
        e.GUID = GUID
        e.Files.Add WSFFile
        AuditedClsIDs.Add e, GUID
    Else
        Set e = AuditedClsIDs(GUID)
        e.Files.Add WSFFile
    End If
    
End Sub

