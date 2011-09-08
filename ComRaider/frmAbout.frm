VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About COMRaider"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   7515
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   1500
      TabIndex        =   0
      Top             =   180
      Width           =   6075
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim tmp() As String
    
    Me.Icon = MDIForm1.Icon
    
    push tmp, "COMRaider v" & App.Major & "." & App.Minor & "." & App.Revision
    push tmp, "Developer: David Zimmer (david@idefense.com)"
    push tmp, "Copyright: iDefense, A Verisign Company"
    push tmp, "Site:         http://labs.idefense.com"
    
    Label1 = Join(tmp, vbCrLf)
    
    Text1 = "COMRaider is a research tool developed to experiment with some ideas on COM Object " & _
            "fuzzing and distributed auditing techniques." & vbCrLf & _
            "" & vbCrLf & _
            "COMRaider is published under GPL License and full source is installed along with the " & _
            "application. Links to the source project files can be found on the start menu. " & _
            "" & vbCrLf & vbCrLf & _
            "COMRaider has been developed in Visual Basic  6,  Visual C++ 6, and uses a local Access 97 database" & _
            " for data storage. It is recommended to use MySql for distributed auditing mode, however " & _
            "any database server should work equally well." & vbCrLf & _
            "" & vbCrLf & _
            "Assembler and Disassembler engines used in hooker.lib and olly.dll are " & _
            "Copyright (C) 2001 Oleh Yuschuk and used under GPL License." & _
            " (disasm.h, asmserv.c, assembl.c, disasm.c). "
            
    
End Sub
