VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEditAuditNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Audit Notes for Clsid"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1800
      Width           =   7875
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   3096
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
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Auditor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEditAuditNotes"
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

Sub EditNotesFor(clsid As String)
    
    Dim rs As Recordset
    Dim li As ListItem
    
    Set rs = QueryDsn("Select * from tblauditlog where clsid='" & clsid & "'")
    
    While Not rs.EOF
        Set li = lv.ListItems.Add(, , rs!sdate)
        li.SubItems(1) = rs!auditor
        li.SubItems(2) = rs!notes
        li.Tag = rs!autoid
        rs.MoveNext
    Wend
    
    Me.Show
    
End Sub

Private Sub cmdUpdate_Click()
    
    If selli Is Nothing Then
        MsgBox "No item selected", vbInformation
        Exit Sub
    End If
    
    Dim sql As String, Data As String
    
    Data = Replace(Text1, "'", "")
    selli.SubItems(2) = Data
    sql = "Update tblauditlog set notes='" & Data & "' where autoid=" & selli.Tag
    QueryDsn sql
   
    MsgBox "Updated", vbInformation
    
End Sub

Private Sub Form_Load()
    SizeLV lv
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
    Text1 = Item.SubItems(2)
End Sub
