VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDistNotes 
   Caption         =   "View Distributed Notes"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   9645
   Begin VB.CommandButton cmdAddNote 
      Caption         =   "Add Note"
      Height          =   315
      Left            =   8460
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   5700
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.ComboBox cboCatagories 
      Height          =   315
      Left            =   4860
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Catagory"
      Height          =   255
      Index           =   3
      Left            =   3660
      TabIndex        =   8
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   7080
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Notes"
      Height          =   255
      Index           =   2
      Left            =   2820
      TabIndex        =   5
      Top             =   0
      Width           =   915
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   2160
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
      NumItems        =   4
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
         Text            =   "Catagory"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
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
      End
   End
End
Attribute VB_Name = "frmDistNotes"
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
Dim viewOnly As Boolean

Private Const cats = "How-to,Bug!,Wish List,Work Around,Targets,Help,Other"

Private Sub cmdAddNote_Click()
    frmAddDistNote.LoadUp cats
End Sub

Sub Command1_Click()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim c As String
    c = Replace(txtSearch, "'", "")
    
    sql = "select * from tbldistnotes"

    Select Case selopt
        Case 1: sql = sql & " where sdate like '%" & c & "%'"
        Case 2: sql = sql & " where notes like '%" & c & "%'"
        Case 3: sql = sql & " where catagory='" & cboCatagories.Text & "'"
        Case 4: sql = sql & " where auditor like '%" & c & "%'"
    End Select
    
    Filllv sql & " order by autoid asc"
        
End Sub

Private Sub Form_Load()

    Dim rs As Recordset
    Dim li As ListItem
    Dim tmp() As String
    Dim x
    
    With cboCatagories
        txtSearch.Move .Left, .top, .Width, .Height
        tmp() = Split(cats, ",")
        For Each x In tmp
            .AddItem x
        Next
        .ListIndex = 0
    End With
    
    
    Me.Icon = MDIForm1.Icon
    Me.Move 0, 0
    SizeLV lv
    
    Filllv "Select * from tbldistnotes"

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
        li.SubItems(2) = rs!catagory
        li.SubItems(3) = rs!notes
        rs.MoveNext
    Wend
    
End Sub
 

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 9765 Then Me.Width = 9765
    lv.Width = Me.Width - 200
    lv.Height = Me.Height - lv.top - 400
    SizeLV lv
End Sub

Private Sub lv_DblClick()
    viewOnly = True
    mnuUpdateNote_Click
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

 

Private Sub mnuDeleteEntry_Click()

    If MsgBox("Are you sure you want to delete all selected entries?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If

    Dim li As ListItem
    For Each li In lv.ListItems
        If li.Selected Then
            QueryDsn "Delete from tbldistnotes where autoid=" & li.Tag
        End If
    Next

    Filllv "Select * from tbldistnotes"

End Sub

Private Sub mnuUpdateNote_Click()
    
'    Dim n As String
'
'    If selli Is Nothing Then
'        MsgBox "No items are selected", vbInformation
'        Exit Sub
'    End If
'
'    n = Replace(frmMsg.GetData("Update Note", selli.SubItems(3)), "'", Empty)
'    If Len(n) = 0 Then Exit Sub
'
'    QueryDsn "Update tbldistnotes set notes='" & n & "' where autoid=" & selli.Tag
'    selli.SubItems(3) = n

    If selli Is Nothing Then Exit Sub
    
    frmAddDistNote.LoadUp cats, selli.Tag, selli.SubItems(3), selli.SubItems(2), viewOnly
    viewOnly = False

    
End Sub

Private Sub optAll_Click(Index As Integer)
    selopt = Index
    cboCatagories.Visible = (Index = 3)
    txtSearch.Visible = Not cboCatagories.Visible
End Sub



