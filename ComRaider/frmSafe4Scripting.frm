VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSafeForScripting 
   Caption         =   "Enumerate Classes Marked as Safe for Scripting / Programmable and Insertable"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   10380
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   3540
      Width           =   10335
      Begin VB.TextBox txtSafetyReport 
         Height          =   1215
         Left            =   6660
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   4620
         TabIndex        =   23
         Top             =   900
         Width           =   3975
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   375
         Left            =   8700
         TabIndex        =   22
         Top             =   900
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search Tools "
         Height          =   855
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10335
         Begin VB.OptionButton optFile 
            Caption         =   "File Name"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optDesc 
            Caption         =   "Description"
            Height          =   255
            Left            =   2280
            TabIndex        =   18
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton optGuid 
            Caption         =   "Guid"
            Height          =   255
            Left            =   3960
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optProgID 
            Caption         =   "ProgId"
            Height          =   255
            Left            =   3960
            TabIndex        =   16
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton optAll 
            Caption         =   "Show Active"
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   7200
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   255
            Left            =   9120
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optDate 
            Caption         =   "Date"
            Height          =   315
            Left            =   5340
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optHidden 
            Caption         =   "Show Hidden"
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   540
            Width           =   1275
         End
         Begin VB.OptionButton optHighlighted 
            Caption         =   "Highlighted"
            Height          =   195
            Left            =   5340
            TabIndex        =   10
            Top             =   600
            Width           =   1155
         End
         Begin VB.OptionButton optAudited 
            Caption         =   "Audited"
            Height          =   195
            Left            =   7200
            TabIndex        =   9
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Search :"
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Contains"
            Height          =   255
            Left            =   6360
            TabIndex        =   20
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.TextBox txtProgID 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox txtDesc 
         Height          =   255
         Left            =   1260
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1620
         Width           =   7335
      End
      Begin VB.TextBox txtGUID 
         Height          =   285
         Left            =   4620
         TabIndex        =   5
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   900
         Width           =   2055
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         Height          =   375
         Left            =   9900
         TabIndex        =   3
         Top             =   900
         Width           =   435
      End
      Begin VB.TextBox txtAuditNotes 
         Height          =   1155
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Label lblKillBitted 
         Caption         =   "Kill Bit is Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Safety Report"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Selected File"
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "ProgID"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "GUID"
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   26
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Date Added"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Audit Notes &&"
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   1980
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
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
         Text            =   "Date"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "GUID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ProgID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Server"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuScanNew 
         Caption         =   "Scan for New"
      End
      Begin VB.Menu mnuHighLight 
         Caption         =   "Highlight Entry"
      End
      Begin VB.Menu mnuViewFileProps 
         Caption         =   "View File Properties"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewObjSafetyReport 
         Caption         =   "View Object Safety Report"
      End
      Begin VB.Menu mnuRemoveNonFuzzables 
         Caption         =   "Remove non-fuzzables"
      End
      Begin VB.Menu mnuStringScanner 
         Caption         =   "Scan Selected For Strings"
      End
      Begin VB.Menu mnuRebuildDB 
         Caption         =   "Rebuild Db from Scratch"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFuzzSelected 
         Caption         =   "Fuzz Selected"
      End
      Begin VB.Menu mnuDeleteSelected 
         Caption         =   "Hide Selected"
      End
      Begin VB.Menu mnuHideAudited 
         Caption         =   "Hide Audited"
      End
      Begin VB.Menu mnuMarkAudited 
         Caption         =   "Mark As Audited"
      End
      Begin VB.Menu mnuAddAuditNote 
         Caption         =   "Add Audit Note"
      End
      Begin VB.Menu mnuViewAuditHistory 
         Caption         =   "View Audit History"
      End
      Begin VB.Menu mnuLoadDistAuditList 
         Caption         =   "Update Distributed Audit List"
      End
   End
End
Attribute VB_Name = "frmSafeForScripting"
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

'Safe for Scripting document:
'http://msdn.microsoft.com/workshop/components/com/iobjectsafetyextensions.asp?frame=true

Option Explicit

Dim selli As ListItem
Dim abort As Boolean

Public WaitingForRemoteRefresh As Boolean
 
Private Sub cmdHelp_Click()

    MsgBox "Couple notes.." & vbCrLf & _
            "" & vbCrLf & _
            "ActiveX Servers can host multiple classes in the same dll file." & vbCrLf & _
            "" & vbCrLf & _
            "Classes are marked as safe for scripting on a class by class basis." & vbCrLf & _
            "To avoid confusion, when you select a class from this interface, the " & vbCrLf & _
            "initial display results will be filtered to only show the Safe for Scripting " & vbCrLf & _
            "class which you have selected." & vbCrLf & _
            "" & vbCrLf & _
            "If you wish to see the rest of the classes in the library, toggle the " & vbCrLf & _
            """show only fuzzable"" checkbox on the tlb viewer form." & vbCrLf & _
            "" & vbCrLf
            
End Sub
 

Private Sub Form_Terminate()
    abort = True
End Sub

 

Private Sub mnuFuzzSelected_Click()
    Dim li As ListItem
    Dim i As Long
    Dim tmp() As String
    Dim match As String
    Dim alerted As New Collection
    
    On Error Resume Next
      
    Const base As String = "C:\COMRaider\AuditList"
      
    If fso.FolderExists(base) Then
        fso.DeleteFolder base, True
        fso.buildPath base
    End If
      
    For i = lv.ListItems.Count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.Selected Then
            If frmtlbViewer.LoadFile(li.SubItems(3), li.SubItems(1), True, False) Then
                    Set frmtlbViewer.ActiveNode = frmtlbViewer.tv.Nodes(1)
                    frmtlbViewer.GenFiles "C:\COMRaider\AuditList"
            End If
            Unload frmtlbViewer
        End If
    Next
                      
    tmp() = fso.GetFolderFiles(base, , , True)
    
    If AryIsEmpty(tmp) Then
        MsgBox "No fuzz files were created for these classes", vbInformation
    Else
        frmCrashMon.LoadFileList tmp, True
    End If
    
End Sub

Private Sub mnuStringScanner_Click()
    Dim li As ListItem
    Dim i As Long
    Dim tmp() As String
    Dim match As String
    Dim alerted As New Collection
    
    On Error Resume Next
    
    match = InputBox("Enter comma delimited substrings to find", , "file,path,url,key")
    If Len(match) = 0 Then Exit Sub
    
    For i = lv.ListItems.Count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.Selected Then
            If frmtlbViewer.LoadFile(li.SubItems(3), li.SubItems(1), True, False) Then
                    frmtlbViewer.ScanElementsFor match, tmp, alerted
            End If
            Unload frmtlbViewer
        End If
    Next
                                     
    If Not AryIsEmpty(tmp) Then
        frmMsg.Display "Search results for match string: " & match & vbCrLf & vbCrLf & Join(tmp, vbCrLf)
    Else
        MsgBox "no string matchs found for function names or arguments :(", vbInformation
    End If
                    
End Sub

Private Sub mnuViewFileProps_Click()
    Dim p As String
    p = ExpandPath(txtFile)
    If fso.FileExists(p) Then frmMsg.Display QuickInfo(p)
End Sub



Private Sub Form_Resize()
    On Error Resume Next
 
    If Me.Width < 10500 Then
        Me.Width = 10500
    Else
        lv.Width = Me.Width - lv.Left - 150
    End If
        
    If Me.Height < 5730 Then
        Me.Height = 5730
    Else
        Frame2.top = Me.Height - Frame2.Height - 350
        lv.Height = Frame2.top - lv.top - 100
    End If
    
    SizeLV lv
    
End Sub

Private Sub lv_DblClick()
    mnuViewAuditHistory_Click
End Sub

 

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then mnuDeleteSelected_Click
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuAddAuditNote_Click()

    If selli Is Nothing Then Exit Sub
    
    If selli.ForeColor <> AuditColor Then
        MsgBox "This item has not been audited!", vbInformation
        Exit Sub
    End If
    
    EditAuditNote selli.SubItems(1)
    
End Sub

Private Sub mnuHideAudited_Click()
    Dim i As Long
    
    For i = lv.ListItems.Count To 1 Step -1
        If lv.ListItems(i).ForeColor = AuditColor Then
            cn.Execute "Update tblGuids set hidden=1 where autoid=" & lv.ListItems(i).Tag
            lv.ListItems.Remove i
        End If
    Next
    
End Sub

Private Sub mnuHighLight_Click()
    Dim li As ListItem
    Dim v As Long
    
    On Error Resume Next
    For Each li In lv.ListItems
        If li.Selected Then
            v = cn.Execute("Select highlight from tblGUIDs where autoid=" & li.Tag)!highlight
            v = IIf(v = 0, 1, 0) 'invert
            cn.Execute "Update tblGUIDs set highlight=" & v & " where autoid=" & li.Tag
            HiLightItem li, v
        End If
    Next
    
End Sub

Private Sub HiLightItem(li As ListItem, Optional lite As Long = 1, Optional color As ColorConstants)
    Dim X
    Dim c As ColorConstants
    
    If color = 0 Then c = vbBlue Else c = color
    If lite = 0 Then c = vbBlack
    
    li.ForeColor = c
    For Each X In li.ListSubItems
        X.ForeColor = c
    Next
    
End Sub

Private Sub mnuLoadDistAuditList_Click()
   UpdateAuditList False, True
End Sub

Sub UpdateAuditList(Optional silentMode As Boolean = False, Optional Display As Boolean = False)
    Dim rs As Recordset
    Dim cnt As Long
    Dim found As Long
       
    Set rs = QueryDsn("Select distinct clsid from tblauditlog") ' where autoid > " & Options.DistributedModeLastAutoid & " order by autoid asc")

    While Not rs.EOF
        cnt = cn.Execute("Select count(autoid) as cnt from tblGUIDs where audited=0 and clsid='" & rs!clsid & "'")!cnt
        If cnt > 0 Then
            found = found + 1
            cn.Execute "Update tblGUIDs set audited=1 where clsid='" & rs!clsid & "'"
        End If
        'Options.DistributedModeLastAutoid = rs!autoid
        rs.MoveNext
    Wend

    If found > 0 Then
        If Display Then Filllv "Select * from tblGUIDs where hidden=0 order by autoid asc"
        If Not silentMode Then MsgBox "Added another " & found & " Clsids to the audited list", vbInformation
    End If
  
        
End Sub

Private Sub mnuMarkAudited_Click()
    
    If selli Is Nothing Then
        MsgBox "There are no selected items", vbInformation
        Exit Sub
    End If
    
    Dim note As String, sql As String, version As String, li As ListItem
    
    note = InputBox("Enter reason you are marking as audited with no tests", , "No fuzzable members")
    note = Replace(note, "'", Empty)
    If Len(note) = 0 Then Exit Sub
    
    For Each li In lv.ListItems
        If li.Selected Then
        
            version = FileVersion(li.SubItems(3))
            
            sql = "Update tblGuids set audited=1 where clsid='" & li.SubItems(1) & "'"
            cn.Execute sql
            
            DsnInsert "tblauditlog", "clsid,progid,auditor,sdate,version,crashs,tests,notes", _
                            li.SubItems(1), li.SubItems(2), Options.UserName, _
                            Format(Now, "m.d.yy"), version, 0, 0, note
                            
            HiLightItem li, , AuditColor
            
        End If
    Next
    
End Sub

Private Sub mnuRebuildDB_Click()
    If lv.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to delete the current database and build a new set?", vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    ScanRegistry
End Sub


Private Sub mnuRemoveNonFuzzables_Click()
    Dim li As ListItem
    Dim i As Long, cnt As Long
    Dim version As String
    Const note As String = "no fuzzable members (autotest)"

    On Error Resume Next
    
    For i = lv.ListItems.Count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.ForeColor <> AuditColor Then
            If Not frmtlbViewer.LoadFile(li.SubItems(3), li.SubItems(1), True, False) Then
                If frmtlbViewer.tlb.NO_VALID_TLB Then
                    'still may be fuzzabled if they choose to live load it..
                Else
                    cnt = cnt + 1
                    cn.Execute "Update tblGUIDs set hidden=1 , audited=1 where clsid='" & li.SubItems(1) & "'"
                    
                    If Options.DistributedMode Then
                        version = FileVersion(li.SubItems(3))
                       
                        DsnInsert "tblauditlog", "clsid,progid,auditor,sdate,version,crashs,tests,notes", _
                            li.SubItems(1), li.SubItems(2), Options.UserName, _
                            Format(Now, "m.d.yy"), version, 0, 0, note
                            
                    End If
                    
                    lv.ListItems.Remove i
                End If
                Unload frmtlbViewer
            End If
        End If
    Next
                                     
    MsgBox cnt & " records found not to be fuzzable", vbInformation
                    
End Sub

Private Sub mnuScanNew_Click()
    ScanRegistry True
End Sub

Sub ScanRegistry(Optional onlyNew As Boolean = False)

    Dim l As String
    
    If Options.UseSimpleRegScanner Then
        SimpleScanRegistry onlyNew
    Else
    
        l = App.path & IIf(IsIde, "\..\", "") & "\builddb.exe"
        If Not fso.FileExists(l) Then
            MsgBox "Could not find: " & l, vbInformation
            Exit Sub
        End If
        
        lv.ListItems.Clear
        If Not onlyNew Then
            cn.Execute "Delete from tblGUIDs"
            cn.Execute "Delete from tblScanned"
        End If
               
        On Error Resume Next
        WaitingForRemoteRefresh = True
        frmAdvScan.Show 1
         
    End If
    
End Sub






Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub mnuDeleteSelected_Click()
    
    Dim i As Long
    Dim X As Long
    Dim li As ListItem
    
    For i = lv.ListItems.Count To 1 Step -1
        If lv.ListItems(i).Selected Then
            cn.Execute "Update tblGuids set hidden=1 where autoid=" & lv.ListItems(i).Tag
            lv.ListItems.Remove i
        End If
    Next

End Sub

Sub cmdSearch_Click()
    
    Dim sql As String
    Dim arg As String
    Dim rs As Recordset
    
    WaitingForRemoteRefresh = False
    
    sql = "Select * from tblGUIDs"
    arg = Replace(txtSearch, "'", Empty)
    
    If Not optAll And Not optHidden And Not optHighlighted And Not optAudited Then
        sql = sql & " where _____ like '%" & arg & "%' order by autoid asc"
    End If
    
    If optFile Then
        sql = Replace(sql, "_____", "InProcServer32")
    ElseIf optDesc Then
        sql = Replace(sql, "_____", "description")
    ElseIf optGuid Then
        sql = Replace(sql, "_____", "clsid")
    ElseIf optProgID Then
        sql = Replace(sql, "_____", "ProgID")
    ElseIf optDate Then
        sql = Replace(sql, "_____", "sDate")
    ElseIf optHidden Then
        sql = sql & " where hidden=1 order by autoid asc"
    ElseIf optHighlighted Then
        sql = sql & " where highlight=1 order by autoid asc"
    ElseIf optAll Then
        sql = sql & " where hidden=0 order by autoid asc"
    ElseIf optAudited Then
        sql = sql & " where audited=1 order by autoid asc"
    End If
    
    Filllv sql
    
End Sub

Private Sub cmdSelect_Click()
    Dim tmp As String
    Dim env As String
    Dim X As Long
    
    txtFile = ExpandPath(txtFile)
        
    If Not fso.FileExists(txtFile) Then
        MsgBox "File not found: " & txtFile
    Else
        If Not frmtlbViewer.LoadFile(txtFile, txtGUID) Then
            'mnuDeleteSelected_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()

     Me.top = 0
     Me.Left = 0
     Me.Icon = MDIForm1.Icon
     
     mnuLoadDistAuditList.Visible = Options.DistributedMode
     optAudited.Visible = Options.DistributedMode
     mnuViewAuditHistory.Visible = Options.DistributedMode
     mnuAddAuditNote.Visible = Options.DistributedMode
     mnuMarkAudited.Visible = Options.DistributedMode
     
    With lv
        .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - 200
    End With
     
    If cn.Execute("Select count(autoid) as cnt from tblGUIDs")!cnt > 0 Then
        If Options.DistributedMode Then UpdateAuditList True, False
        Filllv "Select * from tblGUIDs where hidden=0 order by autoid asc"
    Else
        Me.Visible = True
        If MsgBox("Would you like to build a ""Safe for Scripting"" " & vbCrLf & "database for your machine?", vbYesNo) = vbYes Then
            mnuRebuildDB_Click
        End If
        If Options.DistributedMode Then UpdateAuditList True, True
    End If
            
    
   

End Sub

Sub MarkGUIDAsAudited(clsid As String)
    Dim li As ListItem
    Dim l As ListSubItem
    Dim i As Long
    
    For Each li In lv.ListItems
        If LCase(li.SubItems(1)) = LCase(clsid) Then
            li.ForeColor = AuditColor
            For Each l In li.ListSubItems
                l.ForeColor = AuditColor
            Next
        End If
    Next
                    
End Sub

Sub Filllv(sql As String)
    Dim rs As Recordset
    Dim li As ListItem
    On Error GoTo hell
    
    lv.ListItems.Clear
    Set rs = cn.Execute(sql)
    
    On Error Resume Next
    
    While Not rs.EOF
        Set li = lv.ListItems.Add(, , rs!sdate)
        li.Tag = rs!autoid
        li.SubItems(1) = rs!clsid
        li.SubItems(2) = rs!ProgID
        
        If Len(rs!inprocserver32) > 0 Then
            li.SubItems(3) = rs!inprocserver32
        ElseIf Len(rs!inprochandler32) > 0 Then
            li.SubItems(3) = rs!inprochandler32
        ElseIf Len(rs!localserver32) > 0 Then
            li.SubItems(3) = rs!localserver32
        End If
        
        li.SubItems(4) = rs!Description
        If rs!highlight = 1 Then HiLightItem li
        If rs!audited = 1 Then HiLightItem li, , AuditColor
        
        If KeyExistsInCollection(killbitted, li.SubItems(1)) Then
            li.ForeColor = vbRed
        End If
         
        rs.MoveNext
    Wend
    
    Me.caption = lv.ListItems.Count & " classes returned"
        
    Exit Sub
hell:     MsgBox Err.Description
End Sub

Sub Insert(tblName, fields, ParamArray params())
    Dim sql As String, i As Integer, values(), tn As String
    
    values() = params() 'force byval
    
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "empty" Then
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next

    sql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    sql = Replace(sql, "____", Join(values, ","))
    cn.Execute sql
    
End Sub

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tmp As String
    
    Set selli = Item
    txtDate = Item.Text
    txtGUID = Item.SubItems(1)
    txtProgID = Item.SubItems(2)
    txtFile = Item.SubItems(3)
    txtDesc = Item.SubItems(4)
    txtAuditNotes = Empty
    
    Dim s As Long
    Dim li As ListItem
    Dim safetyReport As String
    
    If GetSavedSafetyReport(txtGUID, safetyReport) Then
        txtSafetyReport = safetyReport
        'mnuViewObjSafetyReport.Visible = False
    Else
        mnuViewObjSafetyReport.Visible = True
        txtSafetyReport = "Information not in database"
    End If
    
    If Options.DistributedMode Then
        For Each li In lv.ListItems
            If li.Selected Then
                s = s + 1
                If s > 1 Then Exit Sub
            End If
        Next
        If GetAuditHistory(Item.SubItems(1), tmp) Then txtAuditNotes = tmp
    End If
    
    lblKillBitted.Visible = KeyExistsInCollection(killbitted, Item.SubItems(1))
        
    
End Sub
 


Private Sub mnuViewAuditHistory_Click()
    
    If Len(txtGUID) = 0 Then Exit Sub
    
    Dim tmp  As String
    If Not GetAuditHistory(selli.SubItems(1), tmp) Then
        MsgBox "This item has not been audited", vbInformation
    Else
        frmAddDistNote.DisplayText tmp, "Viewing Audit History"
    End If
    
    
End Sub

Private Function GetAuditHistory(clsid As String, outval As String) As Boolean

    Dim tmp() As String
    Dim rs As Recordset
    Dim n As String
    
    If selli.ForeColor <> AuditColor Then
        Exit Function
    Else
        On Error Resume Next
        Set rs = QueryDsn("Select * from tblauditlog where clsid='" & clsid & "'")
        
        'push tmp, "Audit report for Clsid: " & clsid
        'push tmp, String(40, "-")
        While Not rs.EOF
           n = IIf(IsNull(rs!notes), "", rs!notes)
           push tmp, "Auditor: " & rs!auditor & vbTab & "Date: " & rs!sdate & vbTab & _
                     "Crashs: " & rs!crashs & vbTab & "Tests: " & rs!tests & vbCrLf & _
                     "Notes: " & n & vbCrLf & String(50, "-") '& vbCrLf
                     
           rs.MoveNext
        Wend
        
        outval = Join(tmp, vbCrLf)
        GetAuditHistory = True
        
    End If
End Function

Private Sub mnuViewObjSafetyReport_Click()
    If selli Is Nothing Then Exit Sub
    ShowObjSafetyReport selli.SubItems(1)
End Sub




Sub SimpleScanRegistry(Optional onlyNew As Boolean = False)

    Dim tmp() As String
    Dim ret()
    Dim v
    Dim i As Long
    Dim handler As String, Server As String
    Dim cnt As Long
    Dim f As String
    
    abort = False
    lv.ListItems.Clear
    If Not onlyNew Then cn.Execute "Delete from tblGUIDs"

    Const SafeForScrCat = "\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}"
    Const SafeForInitCat = "\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}"

    reg.hive = HKEY_CLASSES_ROOT
    tmp() = reg.EnumKeys("\CLSID\")

    Dim ProgID, sdate, desc, inproc, ver
    Dim safeForScript As Boolean
    Dim safeForInit As Boolean
    
    For Each v In tmp
        
        i = i + 1
        Me.caption = "Scanning: " & i & "/" & UBound(tmp)
        DoEvents
        Me.Refresh
        
        If onlyNew Then
            cnt = cn.Execute("Select count(autoid) as cnt from tblscanned where clsid='" & v & "'")!cnt
            If cnt > 0 Then GoTo nextOne
        End If

        'if its a class thats has regkey safe for scripting or init then add
        
        safeForScript = reg.keyExists("\CLSID\" & v & SafeForScrCat)
        safeForInit = reg.keyExists("\CLSID\" & v & SafeForInitCat)
        
        If safeForScript Or safeForInit Then

            f = Empty
            ver = Empty
            sdate = Format(Now, "m.d.yy")
            desc = reg.ReadValue("\CLSID\" & v, "") 'description
            ProgID = reg.ReadValue("\CLSID\" & v & "\ProgID", "")
            inproc = reg.ReadValue("\CLSID\" & v & "\InProcServer32", "")
            handler = reg.ReadValue("\CLSID\" & v & "\InProcHandler32", "")
            Server = reg.ReadValue("\CLSID\" & v & "\Localserver32", "")
            
            If Len(inproc) > 0 Then
                f = inproc
            ElseIf Len(handler) > 0 Then
                f = handler
            ElseIf Len(Server) > 0 Then
                f = Server
            End If
            
            If Len(f) > 0 And fso.FileExists(f) Then ver = FileVersion(f)
                        
            Const fields = "clsid,description,ProgID,InProcHandler32," & _
                           "Localserver32,InProcServer32,sDate,version," & _
                           "safeForScript,safeForInit"
                           
            Insert "tblGUIDs", fields, v, desc, ProgID, handler, _
                               Server, inproc, sdate, ver, _
                               IIf(safeForScript, 1, 0), IIf(safeForInit, 1, 0)
                               
            Insert "tblScanned", "clsid", v

        End If
nextOne:
        
        If abort Then Exit Sub
        DoEvents
    Next

    Me.caption = "Loading results..."
    Me.Refresh
    DoEvents
    
    Filllv "Select * from tblGUIDs where hidden=0"
    Me.caption = lv.ListItems.Count & " classes marked as safe for scripting"

End Sub



