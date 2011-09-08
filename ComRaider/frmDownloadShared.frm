VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDownloadShared 
   Caption         =   "Download Shared Audits"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9645
   Begin VB.OptionButton optAll 
      Caption         =   "All"
      Height          =   255
      Index           =   5
      Left            =   2340
      TabIndex        =   12
      Top             =   60
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   9615
      Begin VB.TextBox txtApiLog 
         Height          =   1455
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   1500
         Width           =   9615
      End
      Begin MSComctlLib.ListView lvMsg 
         Height          =   1455
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   2566
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
         NumItems        =   3
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
            Text            =   "Instruction"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   8520
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   6060
      TabIndex        =   7
      Top             =   60
      Width           =   2415
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Notes"
      Height          =   255
      Index           =   2
      Left            =   5340
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Clsid"
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "User"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Unprocessed"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5212
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Auditor"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ProgID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clsid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Exceptions"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search by"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuViewTlb 
         Caption         =   "View Tlb"
      End
      Begin VB.Menu mnuViewFile 
         Caption         =   "View File"
      End
      Begin VB.Menu mnuUpdateNote 
         Caption         =   "Update Note"
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteEntry 
         Caption         =   "Delete Entry"
      End
      Begin VB.Menu mnuMarkProcessed 
         Caption         =   "Mark as Processed"
      End
      Begin VB.Menu mnuDownloadNFuzz 
         Caption         =   "Download and Fuzz"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuViewExceptionDetails 
         Caption         =   "View Exception Details"
      End
   End
End
Attribute VB_Name = "frmDownloadShared"
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

Dim selli As ListItem
Dim selMsg As ListItem
Dim selopt As Integer

Private Sub Command1_Click()
    'all dat notes, clsid, user
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim c As String
    c = Replace(txtSearch, "'", "")
    
    sql = "select * from tblscripts"

    Select Case selopt
        Case 0: sql = sql & " where processed=0"
        Case 1: sql = sql & " where sdate like '%" & c & "%'"
        Case 2: sql = sql & " where notes like '%" & c & "%'"
        Case 3: sql = sql & " where clsid like '%" & c & "%'"
        Case 4: sql = sql & " where auditor like '%" & c & "%'"
    End Select
    
    Filllv sql
        
End Sub

Private Sub Form_Load()

    Dim rs As Recordset
    Dim li As ListItem
        
    Me.Move 0, 0
    Me.Icon = MDIForm1.Icon
    SizeLV lv
    SizeLV lvMsg
    
    Filllv "Select * from tblscripts where processed=0"
    
    If lv.ListItems.Count = 0 Then
        MsgBox "There are no scripts yet uploaded", vbInformation
    End If

End Sub

Sub Filllv(sql As String)

    On Error Resume Next
    Dim rs As Recordset
    Dim li As ListItem
    Dim d As String
    
    lv.ListItems.Clear
    
    Set rs = QueryDsn(sql)
    If rs Is Nothing Then Exit Sub
    
    While Not rs.EOF
        d = IIf(IsNull(rs!sdate), "[Null]", rs!sdate)
        Set li = lv.ListItems.Add(, , d)
        li.Tag = rs!autoid
        li.SubItems(1) = rs!auditor
        li.ListSubItems(1).Tag = IIf(IsNull(rs!apilog), "", rs!apilog)
        li.SubItems(2) = IIf(IsNull(rs!progid), "", rs!progid)
        li.SubItems(3) = rs!clsid
        li.SubItems(4) = QueryDsn("Select count(autoid) as cnt from tblexceptions where pid=" & rs!autoid)!cnt
        li.SubItems(5) = rs!notes
        If Not IsNull(rs!processed) And rs!processed = 1 Then MarkProcessed li
        rs.MoveNext
    Wend
    
End Sub
 
Sub MarkProcessed(li As ListItem)
    Dim l As ListSubItem
    
    li.ForeColor = vbBlue
    For Each l In li.ListSubItems
        l.ForeColor = vbBlue
    Next
                    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.Width < 9700 Then
        Me.Width = 9700
    Else
        lv.Width = Me.Width - lv.Left - 150
        Frame1.Width = Me.Width
        lvMsg.Width = lv.Width
        txtApiLog.Width = lv.Width
    End If
        
    If Me.Height < 5730 Then
        Me.Height = 5730
    Else
        Frame1.top = Me.Height - Frame1.Height - 350
        lv.Height = Frame1.top - lv.top - 100
    End If
    
    SizeLV lv
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim pid As Long
    Dim rs As Recordset
    Dim li As ListItem
    
    On Error Resume Next
    
    Set selli = Item
    lvMsg.ListItems.Clear
    txtApiLog = Item.ListSubItems(1).Tag 'saved api log
    
    Set rs = QueryDsn("Select * from tblexceptions where pid=" & Item.Tag)
    While Not rs.EOF
        Set li = lvMsg.ListItems.Add(, , Hex(rs!address))
        li.Tag = rs!sdata
        li.SubItems(1) = IIf(IsNull(rs!exception), "", rs!exception)
        li.SubItems(2) = rs!Disasm
        rs.MoveNext
    Wend
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lvMsg_DblClick()
    If selMsg Is Nothing Then Exit Sub
    frmMsg.Display selMsg.Tag, True
End Sub

Private Sub lvMsg_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'text1 = Item.Tag
    Set selMsg = Item
End Sub

Private Sub lvMsg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup2
End Sub

Private Sub mnuDeleteEntry_Click()
    
    If MsgBox("Are you sure you want to delete all selected entries?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim li As ListItem
    For Each li In lv.ListItems
        If li.Selected Then
            QueryDsn "Delete from tblscripts where autoid=" & li.Tag
            QueryDsn "Delete from tblexceptions where pid=" & li.Tag
        End If
    Next
    
    lvMsg.ListItems.Clear
    Filllv "Select * from tblscripts where processed=0"
    
End Sub

Private Sub mnuDownloadNFuzz_Click()
    
    Dim li As ListItem
    Dim tmp() As String
    Dim f As String
    
    Const basedir As String = "C:\Comraider\Downloaded"
    
    fso.DeleteFolder basedir, True
    If Not fso.FolderExists(basedir) Then fso.CreateFolder basedir
    
    For Each li In lv.ListItems
        If li.Selected Then
            f = SafeFreeFileName(basedir, ".wsf")
            fso.WriteFile f, QueryDsn("Select script from tblscripts where autoid=" & li.Tag)!script
            push tmp, f
        End If
    Next
    
    frmCrashMon.LoadFileList tmp, True
    
End Sub

Private Sub mnuMarkProcessed_Click()
    Dim li As ListItem
    
    For Each li In lv.ListItems
        If li.Selected Then
            QueryDsn "Update tblscripts set processed=1 where autoid=" & li.Tag
            MarkProcessed li
        End If
    Next
    
End Sub

Private Sub mnuUpdateNote_Click()
    
    Dim n As String
    
    If selli Is Nothing Then
        MsgBox "No items are selected", vbInformation
        Exit Sub
    End If
    
    n = Replace(frmMsg.GetData("Enter new note text", selli.SubItems(5)), "'", Empty)
    
    If Len(n) = 0 Then Exit Sub
    
    QueryDsn "Update tblscripts set notes='" & n & "' where autoid=" & selli.Tag
    selli.SubItems(5) = n
    
End Sub

Private Sub mnuViewExceptionDetails_Click()
    lvMsg_DblClick
End Sub

Private Sub mnuViewFile_Click()
    If selli Is Nothing Then Exit Sub
    On Error GoTo hell
    frmMsg.Display QueryDsn("Select script from tblscripts where autoid=" & selli.Tag)!script
    Exit Sub
hell:     MsgBox Err.Description
End Sub

Private Sub optAll_Click(Index As Integer)
    selopt = Index
End Sub


Private Sub mnuViewTlb_Click()
    
    If selli Is Nothing Then
        MsgBox "No item selected", vbInformation
        Exit Sub
    End If
    
    Dim tmp() As String
    Dim s As String
    Dim clsid As String

    
    clsid = selli.SubItems(3)
    
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

