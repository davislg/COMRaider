VERSION 5.00
Begin VB.Form frmAddDistNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Distributed Note"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin VB.ComboBox cboCatagories 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Notes"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Catagory"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmAddDistNote"
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

Dim autoid As Long

Private Sub cmdDone_Click()
    Dim tmp As String
    
    If Len(Text1) = 0 Then
        MsgBox "Please enter note text", vbInformation
        Exit Sub
    End If
    
    If autoid = 0 Then
        DsnInsert "tbldistnotes", "auditor,sdate,notes,catagory", _
                    Options.UserName, Format(Now, "m.d.yy"), _
                    Replace(Text1, "'", "''"), cboCatagories.Text
    Else
    
        tmp = Replace(Text1, "'", "''") & vbCrLf & _
              "Last edited by: " & Options.UserName & "  (" & Now & ")"
              
        QueryDsn "Update tbldistnotes set notes='" & tmp & "' , " & _
                  "catagory='" & cboCatagories.Text & "'" & _
                  " where autoid=" & autoid
                  
    End If
    
    frmDistNotes.Command1_Click 'research last search
    Unload Me
                
End Sub

Sub DisplayText(t As String, Optional caption = "")
    
    Label1(0).Visible = False
    Label1(1).Visible = False
    cboCatagories.Visible = False
    Text1.top = 100
    Me.Height = 3060
    Text1.Height = Me.Height - 500
    Text1.Left = 100
    Text1.Width = Me.Width - 230
    Text1 = t
    Me.caption = caption
    Me.Show 1
    
End Sub

Sub LoadUp(cats As String, Optional editAutoID As Long, Optional editText As String, Optional editCat As String, Optional viewOnly As Boolean = False)
    
    autoid = editAutoID
    
    Dim tmp() As String
    Dim x, i As Long
    
    tmp = Split(cats, ",")
    For Each x In tmp
        cboCatagories.AddItem x
        If x = editCat Then i = cboCatagories.ListCount - 1
    Next
    
    If autoid > 0 Then
        Text1 = editText
        cboCatagories.ListIndex = i
        Me.caption = "Edit existing note"
    Else
        cboCatagories.ListIndex = 0
    End If
    
    If viewOnly Then
        Me.Height = 3060
        Me.caption = "View note"
    End If
    
    Me.Show 1
    
End Sub

Private Sub Form_Load()
    Me.Icon = MDIForm1.Icon
End Sub
