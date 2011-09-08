VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Object Safety Report"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   2115
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":030A
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
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
 
Private Sub Form_Load()
  
    Dim pth As String
    
    pth = App.Path & IIf(IsIde, "\..", "") & "\comraider2.mdb"
    If Not fso.FileExists(pth) Then
       pth = App.Path & IIf(IsIde, "\..", "") & "\comraider.mdb"
       If Not fso.FileExists(pth) Then
           MsgBox "Could not find " & pth
           End
       End If
    End If
      
    If cn.State = 0 Then
        cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft " & _
                              "Access Driver (*.mdb)};DBQ=" & pth & ";"
        
         cn.Open
    End If
    
    
    Dim l As String
    Dim isReport As Boolean
    
    l = Replace(Command, """", "")
    
    If InStr(1, l, "/report", vbTextCompare) > 0 Then
        isReport = True
        l = Trim(Replace(l, "/report", ""))
    End If
     
    Dim cc As CClassSafety
    
    If isReport Then
        Set cc = isIOSafe(l)
        Me.Height = 2685
        Text2.Move 0, 0, Me.Width, Me.Height
        Text2.Visible = True
        Text2 = cc.GetReport
        Me.Visible = True
    Else
        Set cc = isIOSafe(l) 'saves to db
        End
    End If
    
   
    
End Sub



Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
