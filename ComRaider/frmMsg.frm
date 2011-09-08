VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   5400
      Width           =   1155
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   9060
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmMsg"
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

Dim waiting As Boolean
Dim Cancel As Boolean
Dim startValue As String
Dim isModal As Boolean

Const extHeight As Long = 6165
Const simpleHeight As Long = 5775
Const staticWidth As Long = 10350

Private Sub cmdCancel_Click()
    Cancel = True
End Sub

Private Sub Form_QueryUnload(DoCancel As Integer, UnloadMode As Integer)
    Cancel = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Width = staticWidth
    Me.Height = IIf(waiting, extHeight, simpleHeight)
End Sub


Sub Display(msg, Optional modal As Boolean = False)
    
    Text1 = msg
    
    If modal Then
        Me.Show 1
    Else
        Me.Visible = True
    End If
    
End Sub

Private Sub cmdDone_Click()
    If isModal Then
        Me.Visible = False
    Else
        waiting = False
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = MDIForm1.Icon
End Sub

Function GetData(Optional caption As String = "Enter data", Optional default As String = Empty)
        
       On Error Resume Next
       
       Text1 = default
       startValue = default
       
       Me.caption = caption
       waiting = True
       Me.Visible = True
       
       If Err.Number = 401 Then 'cant show nonmodal while modal shown
            isModal = True
            Me.Show 1
            waiting = False
       End If
       
       Cancel = False
       
       Do While waiting
            DoEvents
            Sleep 60
            If Cancel Then
                Text1 = startValue
                Exit Do
            End If
       Loop
       
       GetData = Text1
       Unload Me
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    waiting = False
End Sub
