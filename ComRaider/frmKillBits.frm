VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmKillBits 
   Caption         =   "View controls with killbit set"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   10380
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3540
      Width           =   10335
      Begin VB.TextBox txtSafetyReport 
         Height          =   1035
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   780
         Width           =   4635
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   3975
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox txtProgID 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   420
         Width           =   2655
      End
      Begin VB.TextBox txtAuditNotes 
         Height          =   975
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   780
         Width           =   4335
      End
      Begin VB.TextBox txtGUID 
         Height          =   285
         Left            =   4560
         TabIndex        =   2
         Top             =   420
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "Audit Notes &&"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Safety Report"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Audit Notes"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Selected File"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "ProgID"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         Height          =   255
         Index           =   0
         Left            =   5220
         TabIndex        =   8
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "GUID"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   7
         Top             =   420
         Width           =   1215
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
         Text            =   "InProcServer"
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
      Begin VB.Menu mnuMarkAsAudited 
         Caption         =   "Mark Audited"
      End
      Begin VB.Menu mnuViewAuditLog 
         Caption         =   "View Audit Log"
      End
      Begin VB.Menu mnuViewFileProps 
         Caption         =   "View File Properties"
      End
      Begin VB.Menu mnuViewObjSafetyReport 
         Caption         =   "View Object Safety Report"
      End
      Begin VB.Menu mnuBuildObjSafetyReport 
         Caption         =   "Build Obj Safety Report for Selected"
      End
      Begin VB.Menu mnuRemoveNonFuzzables 
         Caption         =   "Remove Non Fuzzables"
      End
      Begin VB.Menu mnuFuzzSelected 
         Caption         =   "Fuzz Selected"
      End
      Begin VB.Menu mnuStringScanner 
         Caption         =   "Scan For Strings"
      End
   End
End
Attribute VB_Name = "frmKillBits"
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
Dim running As Boolean

Sub MarkGUIDAsAudited(clsid As String)

    Dim li As ListItem
    Dim X As ListSubItem

    For Each li In lv.ListItems
        If InStr(1, li.SubItems(1), clsid, vbTextCompare) > 0 Then MarkAudited li
    Next

End Sub

 

Private Sub cmdProperties_Click()
    If fso.FileExists(txtFile) Then
        frmMsg.Display QuickInfo(txtFile)
    End If
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
        Frame1.top = Me.Height - Frame1.Height - 350
        lv.Height = Frame1.top - lv.top - 100
    End If
    
    SizeLV lv
    
    
End Sub

Sub DisplayKillBitList()

    Dim tmp() As String
    Dim ret()
    Dim v
    Dim li As ListItem
    Dim handler As String, Server As String
    Dim cnt As Long
    
    Set selli = Nothing
    lv.ListItems.Clear
    
    Const clsid = "\CLSID\"
    
    reg.hive = HKEY_CLASSES_ROOT
     
    Dim path
    Dim auditcnt As Long
    
    running = True
    
    For Each v In killbitted

         If Not reg.keyExists(clsid & v & "\InProcServer32") Then GoTo nextOne
                
         path = reg.ReadValue(clsid & v & "\InProcServer32", "")
         
         If Len(path) = 0 Then GoTo nextOne
          
         Set li = lv.ListItems.Add(, , Format(Now, "m.d.yy"))
         li.SubItems(1) = v
         li.SubItems(3) = path
         li.SubItems(4) = reg.ReadValue(clsid & v, "") 'description
            
         If reg.keyExists(clsid & v & "\Progid") Then
            li.SubItems(2) = reg.ReadValue(clsid & v & "\ProgID", "")
         End If
         
         If Options.DistributedMode Then
                auditcnt = QueryDsn("Select count(autoid) as cnt from tblauditlog where clsid='" & v & "'")!cnt
                If auditcnt > 0 Then MarkAudited li
         End If
        
nextOne:
        DoEvents
    Next
    
    running = False
    
   
    
    If lv.ListItems.Count > 0 Then
 
        Me.caption = lv.ListItems.Count & " classes have been killbitted on this system"
        Me.Visible = True
         
    Else
        'If autoUnload Then
        '    Unload Me
        'Else
'
 '       End If
    End If
    
    
              
End Sub
 
Sub MarkAudited(li As ListItem)
    Dim l As ListSubItem
    
    li.ForeColor = AuditColor
    For Each l In li.ListSubItems
        l.ForeColor = AuditColor
    Next
                    
End Sub


 

Private Sub cmdSelect_Click()
    txtFile = ExpandPath(txtFile)
    If Not fso.FileExists(txtFile) Then
        MsgBox "File not found: " & txtFile
    Else
        frmtlbViewer.LoadFile txtFile, selli.SubItems(1)
    End If
End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.Left = 0
    Me.Icon = MDIForm1.Icon
    mnuViewAuditLog.Visible = Options.DistributedMode
    mnuMarkAsAudited.Visible = Options.DistributedMode
     
    With lv
        .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - 200
    End With

End Sub


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
    txtGUID = Item.SubItems(1)
    txtProgID = Item.SubItems(2)
    txtFile = Item.SubItems(3)
    txtDesc = Item.SubItems(4)
    
    txtSafetyReport = Empty
    
    On Error Resume Next
    Dim rs As Recordset
    Dim tmp() As String
    Dim safetyReport As String
    
    If GetSavedSafetyReport(txtGUID, safetyReport) Then
        txtSafetyReport = safetyReport
        'mnuViewObjSafetyReport.Visible = False
    Else
        mnuViewObjSafetyReport.Visible = True
        txtSafetyReport = "Information not in database"
    End If
    
    If Options.DistributedMode Then
        Set rs = QueryDsn("Select notes,auditor,sdate from tblauditlog where clsid='" & selli.SubItems(1) & "'")
        While Not rs.EOF
             push tmp, rs!sdate & " - " & rs!auditor & " - " & rs!notes & vbCrLf
             rs.MoveNext
        Wend
        txtAuditNotes = Join(tmp, vbCrLf)
    End If
    
End Sub

Private Sub lv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Dim i As Long
        For i = lv.ListItems.Count To 1 Step -1
            If lv.ListItems(i).Selected Then
                lv.ListItems.Remove i
            End If
        Next
    End If
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuBuildObjSafetyReport_Click()
    
    Dim tmp() As String
    Dim li As ListItem
    
    For Each li In lv.ListItems
        If li.Selected Then
            push tmp, li.SubItems(1)
        End If
    Next
    
    If AryIsEmpty(tmp) Then
        MsgBox "There were no classes selected", vbInformation
    Else
        frmAdvScan.BuildForGuidList tmp
    End If
    
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

Private Sub mnuMarkAsAudited_Click()
    
    If selli Is Nothing Then
        MsgBox "There are no selected items", vbInformation
        Exit Sub
    End If
    
    Dim note As String, sql As String, version As String, li As ListItem
    
    note = InputBox("Enter reason you are marking as audited with no tests", , "No fuzzable members")
    note = Replace(note, "'", Empty)
    If Len(note) = 0 Then
        MsgBox "You must enter a reason", vbInformation
        Exit Sub
    End If
    
    For Each li In lv.ListItems
        If li.Selected Then
        
            version = FileVersion(li.SubItems(3))
            
            sql = "Update tblGuids set audited=1 where clsid='" & li.SubItems(1) & "'"
            cn.Execute sql
            
            DsnInsert "tblauditlog", "clsid,progid,auditor,sdate,version,crashs,tests,notes", _
                            li.SubItems(1), li.SubItems(2), Options.UserName, _
                            Format(Now, "m.d.yy"), version, 0, 0, note
                            
            MarkAudited li
            
        End If
    Next
    
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

Private Sub mnuViewAuditLog_Click()
    
    If selli Is Nothing Then Exit Sub
    If Len(txtGUID) = 0 Then Exit Sub
    
    Dim tmp() As String
    Dim rs As Recordset
    
    If selli.ForeColor <> AuditColor Then
        MsgBox "This item has not been audited", vbInformation
    Else
    
        Set rs = QueryDsn("Select * from tblauditlog where clsid='" & selli.SubItems(1) & "'")
        
        push tmp, "Audit report for Clsid: " & selli.SubItems(1)
        push tmp, String(40, "-")
        While Not rs.EOF
            push tmp, "Auditor: " & rs!auditor & vbTab & "Date: " & rs!sdate & vbTab & _
                     "Crashs: " & rs!crashs & vbTab & "Tests: " & rs!tests & vbCrLf & _
                     "Notes: " & rs!notes & vbCrLf & String(50, "-") & vbCrLf
                     
           rs.MoveNext
        Wend
        
        frmMsg.Display Join(tmp, vbCrLf)
        
    End If
    
End Sub

Private Sub mnuViewFileProps_Click()
    Dim p As String
    p = ExpandPath(txtFile)
    If fso.FileExists(p) Then frmMsg.Display QuickInfo(p)
End Sub

Private Sub mnuViewObjSafetyReport_Click()
    If selli Is Nothing Then Exit Sub
    ShowObjSafetyReport txtGUID
End Sub
