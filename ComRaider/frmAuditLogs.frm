VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAuditLogs 
   Caption         =   "View Distributed Audits Logs"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9645
   Begin VB.OptionButton optAll 
      Caption         =   "Crashs"
      Height          =   255
      Index           =   6
      Left            =   4980
      TabIndex        =   10
      Top             =   0
      Width           =   915
   End
   Begin VB.OptionButton optAll 
      Caption         =   "ProgId"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   8160
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   5880
      TabIndex        =   7
      Top             =   0
      Width           =   2115
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Notes"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   6
      Top             =   0
      Width           =   915
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   3540
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Clsid"
      Height          =   255
      Index           =   3
      Left            =   1980
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "User"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.OptionButton optAll 
      Caption         =   "All"
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
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
      NumItems        =   8
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
         Text            =   "Clsid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Progid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Exceptions"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tests"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search by"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdateNote 
         Caption         =   "Update Note"
      End
      Begin VB.Menu mnuDeleteEntry 
         Caption         =   "Delete Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyClsid 
         Caption         =   "Copy Clsid"
      End
      Begin VB.Menu mnuViewTlb 
         Caption         =   "View Tlb"
      End
   End
End
Attribute VB_Name = "frmAuditLogs"
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
Dim selopt As Integer

Private Sub Command1_Click()
    'all dat notes, clsid, user
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim c As String
    c = Replace(txtSearch, "'", "")
    
    sql = "select * from tblauditlog"

    Select Case selopt
        Case 1: sql = sql & " where sdate like '%" & c & "%'"
        Case 2: sql = sql & " where notes like '%" & c & "%'"
        Case 3: sql = sql & " where clsid like '%" & c & "%'"
        Case 4: sql = sql & " where auditor like '%" & c & "%'"
        Case 5: sql = sql & " where progid like '%" & c & "%'"
        Case 6:
                If Not IsNumeric(c) Then
                    MsgBox "Please enter the min number of crashs to match", vbInformation
                    Exit Sub
                End If
                sql = sql & " where crashs >=" & c
    End Select
    
    Filllv sql & " order by autoid asc"
        
End Sub

Private Sub Form_Load()

    Dim rs As Recordset
    Dim li As ListItem
        
    Me.Icon = MDIForm1.Icon
    SizeLV lv
    
    Filllv "Select * from tblauditlog"
    
    If lv.ListItems.Count = 0 Then
        MsgBox "There are no audits yet uploaded", vbInformation
    End If

End Sub

Sub Filllv(sql As String)

    On Error Resume Next
    Dim rs As Recordset
    Dim li As ListItem
    Dim d As String
    Dim exceptions As Long
    Dim tests As Long
    
    lv.ListItems.Clear
    
    Set rs = QueryDsn(sql)
    If rs Is Nothing Then Exit Sub
    
    While Not rs.EOF
        d = IIf(IsNull(rs!sdate), "[Null]", rs!sdate)
        Set li = lv.ListItems.Add(, , d)
        li.Tag = rs!autoid
        li.SubItems(1) = rs!auditor
        li.SubItems(2) = rs!clsid
        li.SubItems(3) = rs!ProgID
        li.SubItems(4) = rs!crashs
        li.SubItems(5) = rs!tests
        li.SubItems(6) = rs!version
        li.SubItems(7) = rs!notes
        exceptions = exceptions + rs!crashs
        tests = tests + rs!tests
        rs.MoveNext
    Wend
    
    Me.caption = "Searched Audits Encompass " & exceptions & " Exceptions Found in " & tests & " tests for " & lv.ListItems.Count & " COM Classes"
    
End Sub
 

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 9765 Then Me.Width = 9765
    lv.Width = Me.Width - 250
    lv.Height = Me.Height - lv.top - 400
    SizeLV lv
End Sub

Private Sub lv_DblClick()
    mnuUpdateNote_Click
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyClsid_Click()
    If selli Is Nothing Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText selli.SubItems(2)
    MsgBox "Class id copied", vbInformation
End Sub

'Private Sub mnuDeleteEntry_Click()
'
'    If MsgBox("Are you sure you want to delete all selected entries?", vbInformation + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'
'    Dim li As ListItem
'    For Each li In lv.ListItems
'        If li.Selected Then
'            QueryDsn "Delete from tblscripts where autoid=" & li.Tag
'            QueryDsn "Delete from tblexceptions where pid=" & li.Tag
'        End If
'    Next
'
'    lv.ListItems.Clear
'    Form_Load
'
'End Sub

Private Sub mnuUpdateNote_Click()
    
    Dim n As String
    
    If selli Is Nothing Then
        MsgBox "No items are selected", vbInformation
        Exit Sub
    End If
    
    n = Replace(frmMsg.GetData("Update Note", selli.SubItems(7)), "'", Empty)
    If Len(n) = 0 Then Exit Sub
    
    QueryDsn "Update tblauditlog set notes='" & n & "' where autoid=" & selli.Tag
    selli.SubItems(7) = n
    
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

    
    clsid = selli.SubItems(2)
    
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

