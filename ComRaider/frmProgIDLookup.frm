VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgIDLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProgID Lookup"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   315
      Left            =   4020
      TabIndex        =   6
      Top             =   3540
      Width           =   1035
   End
   Begin VB.TextBox txtClsid 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   2940
      Width           =   4395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtProgID 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblCurVer 
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   3300
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Clsid"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   2940
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "ProgID String"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuViewObjSafetyReport 
         Caption         =   "View Obj Safety Report"
      End
   End
End
Attribute VB_Name = "frmProgIDLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim keys() As String
Dim selLi As ListItem

Private Sub cmdSearch_Click()
    
    If Len(txtProgID) = 0 Then
        MsgBox "Please enter a partial ProgID string to match", vbInformation
        Exit Sub
    End If
    
    If Len(txtProgID) < 3 Then
        MsgBox "Please be more specific you would get to many results with this search string", vbInformation
        Exit Sub
    End If
        
    Dim v, tmp
    Dim li As ListItem
    
    lv.ListItems.Clear
    
    tmp = LCase(txtProgID)
    If InStr(tmp, "*") < 1 Then tmp = "*" & tmp & "*"
    
    For Each v In keys
        If LCase(v) Like tmp Then
            lv.ListItems.Add , , v
        End If
    Next
    
End Sub

Private Sub cmdSelect_Click()
    Dim ff As String, f As String
    On Error GoTo hell
    
    ff = txtClsid
    If Len(ff) = 0 Then Exit Sub
        
    reg.hive = HKEY_CLASSES_ROOT
    
    If Right(ff, 1) <> "}" Then ff = ff & "}"
    If Left(ff, 1) <> "{" Then ff = "{" & ff
    
    f = "\CLSID\" & ff
    
    If reg.keyExists(f) Then
        f = f & "\InProcServer32"
        If reg.keyExists(f) Then
            f = reg.ReadValue(f, "")
            f = StripQuotes(f)
            If fso.FileExists(f) Then
                frmtlbViewer.LoadFile f, ff
                Unload Me
            End If
        Else
            MsgBox "Could not find its InProcServer32 entry", vbInformation
        End If
    Else
        MsgBox "Could not locate this GUID on your system", vbInformation
    End If
    
    
    Exit Sub
hell:     MsgBox Err.Description
    
End Sub

Private Sub Form_Load()
    reg.hive = HKEY_CLASSES_ROOT
    keys = reg.EnumKeys("\")
    lv.ColumnHeaders(1).Width = lv.Width
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLi = Item
    
    Dim clsid As String
    Dim curver As String
    
   txtClsid = reg.ReadValue("\" & Item.Text & "\Clsid", "")
   curver = reg.ReadValue("\" & Item.Text & "\CurVer", "")
   
   If Len(curver) > 0 Then
        lblCurVer(1) = "CurVer: " & curver
   Else
        lblCurVer(1) = Empty
   End If
    
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuViewObjSafetyReport_Click()
    If selLi Is Nothing Then Exit Sub
    If Len(txtClsid) = 0 Then
        MsgBox "No clsid found?", vbInformation
        Exit Sub
    End If
    ShowObjSafetyReport txtClsid
End Sub
