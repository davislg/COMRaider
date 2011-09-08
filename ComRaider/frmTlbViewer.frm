VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmtlbViewer 
   BackColor       =   &H8000000A&
   Caption         =   "COMRaider "
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   9900
   Visible         =   0   'False
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5640
      Width           =   9795
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "?"
      Height          =   255
      Left            =   6540
      TabIndex        =   8
      Top             =   120
      Width           =   315
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   2040
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cndNext 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   7215
   End
   Begin VB.CheckBox chkOnlyFuzzable 
      Caption         =   "Show only fuzzable"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   2700
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   480
      Width           =   7215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0000
            Key             =   "const"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0112
            Key             =   "event"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0224
            Key             =   "class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0336
            Key             =   "interface"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0448
            Key             =   "lib"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":055A
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":066C
            Key             =   "module"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":077E
            Key             =   "value"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0890
            Key             =   "prop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8493
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "D:\VCpp\vuln\Debug\vuln.dll"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Initilization Fuzz Script prologue"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "More >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7980
      MouseIcon       =   "frmTlbViewer.frx":09A2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Use right click menu on treeview to generate fuzz files"
      Height          =   435
      Left            =   2700
      TabIndex        =   7
      Top             =   4860
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "COM Server"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuFuzzMember 
         Caption         =   "Fuzz member"
      End
      Begin VB.Menu mnuFuzzMembers 
         Caption         =   "Fuzz Interface"
      End
      Begin VB.Menu mnuFuzzClass 
         Caption         =   "Fuzz Class"
      End
      Begin VB.Menu mnuViewAuditLog 
         Caption         =   "View Audit Log"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddAuditNote 
         Caption         =   "Add Audit Note"
      End
      Begin VB.Menu mnuObjectSafetyReport 
         Caption         =   "ObjectSafety Report"
      End
      Begin VB.Menu mnuTestInIE 
         Caption         =   "Test Scriptable in IE"
      End
      Begin VB.Menu mnuFuzzLibrary 
         Caption         =   "Fuzz Library"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandAll 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuShowAllClasses 
         Caption         =   "Show All Clases"
      End
      Begin VB.Menu mnuLoadFromMem 
         Caption         =   "Load From Memory"
      End
      Begin VB.Menu mnuStringScanner 
         Caption         =   "Scan for Strings"
      End
   End
End
Attribute VB_Name = "frmtlbViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         disassembler functionality provided by olly.dll which
'         is a modified version of the OllyDbg GPL source from
'         Oleh Yuschuk Copyright (C) 2001 - http://ollydbg.de
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

Public tlb As New CTlbParse
Public ActiveNode As Node
Public FilterGUID As String

Public lngs As Collection
Public ints As Collection
Public dbls As Collection
Public strs As Collection
Public cust As Collection
 
Private LiveLoadWarned As Boolean
Private MoreMode As Boolean

Private Sub chkOnlyFuzzable_Click()
    LoadFile Text1, FilterGUID
End Sub

Sub MarkGUIDAsAudited(clsid As String)
    
    Dim c As CClass
    On Error Resume Next
    
    For Each c In tlb.mClasses
        If LCase(c.GUID) = LCase(clsid) Then c.BeenAudited = True
    Next
    
End Sub

Function LoadFile(fPath As String, Optional onlyShowGuid As String, Optional silent As Boolean = False, Optional allowLiveLoad As Boolean = True, Optional forceLiveLoad As Boolean = False) As Boolean
    
    Text1 = ExpandPath(fPath)
    
    Dim c As CClass
    Dim i As CInterface
    Dim m As CMember
    Dim pi As ParameterInfo
    
    Dim n0 As Node
    Dim n1 As Node
    Dim n2 As Node
    Dim n3 As Node
    
    Dim mMembers As Long
    Dim mInterfaces As Long
    Dim x As Long
    Dim loaded As Boolean
    
    Set tlb = Nothing
    Set tlb = New CTlbParse
    
    FilterGUID = onlyShowGuid
    mnuShowAllClasses.Visible = (Len(onlyShowGuid) > 0)
    
    If Len(onlyShowGuid) > 0 Then
        Me.caption = Me.caption & "  Only showing class " & onlyShowGuid
    End If
    
    tv.Nodes.Clear
    List1.Clear
    Text2 = Empty
    
    If Not forceLiveLoad Then
        loaded = tlb.LoadFile(Text1, onlyShowGuid)
    Else
        tlb.NO_VALID_TLB = True
        allowLiveLoad = True
        If Len(onlyShowGuid) = 0 Then MsgBox "improper use of forceLiveLoadFlag no guid!"
    End If
    
    If Not loaded And tlb.NO_VALID_TLB And allowLiveLoad And Len(onlyShowGuid) > 0 Then
        If Not LiveLoadWarned Then
            If MsgBox("Files does not have a valid typelib" & vbCrLf & "Do you want to create a live instance to query?", vbInformation + vbYesNo) = vbYes Then
                LiveLoadWarned = True
            End If
        End If
        If LiveLoadWarned Then
            loaded = tlb.LoadFromMem(onlyShowGuid)
            mnuShowAllClasses.Enabled = False
            lblStatus = lblStatus & vbCrLf & "Loaded class from live instance"
        End If
    End If
                
    If loaded Then
        
        Set n0 = tv.Nodes.Add(, , , tlb.libName, "lib")
        
        For Each c In tlb.mClasses
        
            If Len(onlyShowGuid) > 0 Then
                If InStr(1, c.GUID, onlyShowGuid, vbTextCompare) < 1 Then
                    GoTo nextOne
                End If
            End If
            
            Set n1 = tv.Nodes.Add(n0, tvwChild, , c.Name, "class")
            Set n1.Tag = c
            mInterfaces = 0
            mMembers = 0
            For Each i In c.mInterfaces
                mInterfaces = mInterfaces + 1
                If Options.OnlyDefInf And Not i.isDefault Then GoTo nextOne
                Set n2 = tv.Nodes.Add(n1, tvwChild, , i.Name, "interface")
                Set n2.Tag = i
                For Each m In i.mMembers
                    If m.SupportsFuzzing Or chkOnlyFuzzable.value = 0 Then
                        Set n3 = tv.Nodes.Add(n2, tvwChild, , m.mMemberInfo.Name, IIf(m.CallType > 1, "prop", "sub"))
                        Set n3.Tag = m
                        mMembers = mMembers + 1
                    End If
                    If ObjPtr(n3) And Not m.SupportsFuzzing Then n3.ForeColor = &H606060
                    Set n3 = Nothing
                Next
                n2.Sorted = True
            Next
            If mInterfaces = 0 Or mMembers = 0 Then n1.Tag = Empty
            n1.Sorted = True
nextOne:
        Next
        
        For x = tv.Nodes.Count To 1 Step -1
            If tv.Nodes(x).Index <> n0.Index Then
                If Not IsObject(tv.Nodes(x).Tag) Then
                    tv.Nodes.Remove x
                Else
                    If TypeName(tv.Nodes(x).Tag) = "CInterface" Then
                        If tv.Nodes(x).Children = 0 Then tv.Nodes.Remove x
                    End If
                End If
            End If
        Next
        
        n0.Expanded = True
    
    End If
    
    If tv.Nodes.Count = 0 Then
        If Not silent Then MsgBox tlb.ErrMsg
        LoadFile = False
    ElseIf tv.Nodes.Count = 1 Then
        If chkOnlyFuzzable.value = 1 Then
            chkOnlyFuzzable.value = 0
            LoadFile = False
            If Not silent Then
                LoadFile fPath, onlyShowGuid  'reload with same options
                MsgBox "This " & IIf(Len(onlyShowGuid) = 0, "library", "class") & " did not contain any fuzzable elements :(", vbInformation
            End If
        End If
    Else
        LoadFile = True
        mnuExpandAll_Click
        Me.Visible = True
        Me.ZOrder 0
    End If

End Function


'no longer a real button event just kept as sub name
Sub GenFiles(Optional ByVal pFolder As String = "C:\COMRaider")
    
    On Error Resume Next
    
    Dim c As CMember
    Dim log() As String
    Dim Files() As String
    Dim x
    Dim n As Node
    Dim cnt As Long
    Dim script As String
    
    List1.Clear
    
    If IsIde() Then
        script = App.path & "\..\BuildArgs.vbs"
    Else
         script = App.path & "\BuildArgs.vbs"
    End If
    
    If Not fso.FileExists(script) Then
        MsgBox "Could not load fuzz parameters from: " & script, vbInformation
        Exit Sub
    Else
        With sc
            'reset our fuzz parameters collections to empty
            Set lngs = New Collection
            Set ints = New Collection
            Set dbls = New Collection
            Set strs = New Collection
            Set cust = New Collection
            
            'reload out external vbs script of tests
            .Reset
            .AddObject "parent", Me
            Err.Clear
            .AddCode fso.ReadFile(script)
            
            If Err.Number > 0 Then
                MsgBox "Error adding buildargs.vbs code" & scerr(), vbExclamation
                Err.Clear
                GoTo hell
            End If
            
            'have the script build our argument fuzz list per variable type
            .Eval "GetLongArgs()"
            If Err.Number > 0 Then
                MsgBox "Error in in Buildargs.vbs GetLongArgs" & scerr(), vbExclamation
                Err.Clear
                GoTo hell
            End If
            
            .Eval "GetIntArgs()"
            If Err.Number > 0 Then
                MsgBox "Error in Buildargs.vbs GetIntArgs" & scerr(), vbExclamation
                Err.Clear
                GoTo hell
            End If
            
            .Eval "GetDblArgs()"
            If Err.Number > 0 Then
                MsgBox "Error in Buildargs.vbs GetDblArgs" & scerr(), vbExclamation
                Err.Clear
                GoTo hell
            End If
            
            .Eval "GetStrArgs()"
            If Err.Number > 0 Then
                MsgBox "Error in Buildargs.vbs GetStrArgs" & scerr(), vbExclamation
                Err.Clear
                GoTo hell
            End If
        End With
    End If
    
    
    If TypeName(ActiveNode.Tag) = "CMember" Then
        MakeFuzzFiles ActiveNode, , pFolder
        cnt = 1
    ElseIf TypeName(ActiveNode.Tag) = "CInterface" Then
        For Each n In tv.Nodes
            If Not n.Parent Is Nothing Then
                If n.Parent = ActiveNode Then
                    MakeFuzzFiles n, , pFolder
                    cnt = cnt + 1
                End If
            End If
        Next
    ElseIf TypeName(ActiveNode.Tag) = "CClass" Then
        For Each n In tv.Nodes
            If Not n.Parent Is Nothing Then
                If Not n.Parent.Parent Is Nothing Then
                    If n.Parent.Parent = ActiveNode Then
                        MakeFuzzFiles n, , pFolder
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next
    ElseIf ActiveNode.Index = 1 Then 'root node fuzz whole library
        For Each n In tv.Nodes
            If IsObject(n.Tag) Then
                If TypeName(n.Tag) = "CMember" Then
                    MakeFuzzFiles n, , pFolder
                    cnt = cnt + 1
                End If
            End If
        Next
    End If
        
hell:
    lblStatus = List1.ListCount & " fuzz files created for " & cnt & " functions"
        
End Sub

Sub MakeFuzzFiles(n As Node, Optional overRide As Boolean = False, Optional ByVal pFolder As String = "")
    Dim c As CMember
    Dim x
    Dim log() As String
    Dim Files() As String
    Dim tNode As Node
    Dim proLogue As String
    
    On Error GoTo hell
    
    If Len(pFolder) = 0 Then
        pFolder = "C:\COMRaider"
    Else
        If Not fso.FolderExists(pFolder) Then
            fso.buildPath pFolder
            If Not fso.FolderExists(pFolder) Then
                MsgBox "Could not create path: " & pFolder, vbInformation
                Exit Sub
            End If
        End If
    End If
    
    Set tNode = n.Parent
    If TypeName(tNode.Tag) <> "CClass" Then
        Set tNode = n.Parent.Parent
        If TypeName(tNode.Tag) <> "CClass" Then
            Stop
        Else
            pFolder = pFolder & "\" & tv.Nodes(1).Text & "\" & tNode.Text
        End If
    Else
        pFolder = pFolder & "\" & tv.Nodes(1).Text & "\" & tNode.Text
        Debug.Print "Shouldnt be here"
    End If
    
    If MoreMode And Len(Text3) > 0 Then proLogue = Text3
    
    Set c = n.Tag
    If c.SupportsFuzzing Or overRide Then
        Files() = c.BuildFuzzConfigFiles(log, pFolder, proLogue, tlb.libName, tNode.Text)
        If AryIsEmpty(Files) Then
            Debug.Print "BuildFuzzConfigFiles returned 0 files Text:" & n.Text
        Else
            For Each x In Files
                List1.AddItem x
            Next
            DoEvents
            List1.Refresh
        End If
    End If
    
    
    Exit Sub
hell:     MsgBox "Error in MakeFuzzFiles typename(n.tag)=" & TypeName(n.Tag) & " Desc" & Err.Description
End Sub

Private Sub cmdProperties_Click()
    On Error Resume Next
    If Not fso.FileExists(Text1) Then
        MsgBox "File not found " & Text1
    Else
        frmMsg.Display QuickInfo(Text1)
    End If
End Sub

Private Sub cndNext_Click()

    If List1.ListCount = 0 Then
        MsgBox "You must create fuzz files to proceede", vbInformation
        Exit Sub
    End If
    
    Dim tmp() As String
    Dim i As Long
    
    For i = 0 To List1.ListCount
        push tmp, List1.List(i)
    Next
    
    frmCrashMon.LoadFileList tmp
    
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.Left = 0
    Me.Icon = MDIForm1.Icon
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Height = IIf(MoreMode, 7290, 5715)
    Me.Width = 10020 '8820
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LiveLoadWarned = False
End Sub

Private Sub Label1_Click()
    MoreMode = Not MoreMode
    Label1.caption = IIf(MoreMode, "<< Less", "More >>")
    Form_Resize
End Sub

Private Sub mnuAddAuditNote_Click()
    If ActiveNode Is Nothing Then Exit Sub
    Dim cc As CClass
    
    Set cc = ActiveNode.Tag
    If Not cc.BeenAudited Then
        MsgBox "This class has not yet been audited!", vbInformation
        Exit Sub
    End If
    
    EditAuditNote cc.GUID
    
    MsgBox "Notes Updated", vbInformation
    
End Sub

Private Sub mnuCollapseAll_Click()
    Dim n As Node
    For Each n In tv.Nodes
        If n.Children > 0 Then n.Expanded = False
    Next
    tv.Nodes(1).Expanded = True
    tv.Nodes(1).EnsureVisible
End Sub

Private Sub mnuExpandAll_Click()
    Dim n As Node
    For Each n In tv.Nodes
        If n.Children > 0 Then n.Expanded = True
    Next
    tv.Nodes(1).EnsureVisible
End Sub

Private Sub mnuFuzzClass_Click()
    GenFiles
End Sub

Private Sub mnuFuzzLibrary_Click()
    Set ActiveNode = tv.Nodes(1)
    GenFiles
End Sub

Private Sub mnuFuzzMember_Click()
    GenFiles
End Sub

Private Sub mnuFuzzMembers_Click()
    GenFiles
End Sub

Private Sub mnuLoadFromMem_Click()
    
    If Len(FilterGUID) = 0 Then Exit Sub
    LoadFile "", FilterGUID, , , True
    
End Sub

Private Sub mnuObjectSafetyReport_Click()
    If ActiveNode Is Nothing Then Exit Sub
    If Not IsObject(ActiveNode.Tag) Then Exit Sub
    If TypeName(ActiveNode.Tag) <> "CClass" Then Exit Sub
    Dim c As CClass
    Set c = ActiveNode.Tag
    ShowObjSafetyReport c.GUID
End Sub

Private Sub mnuShowAllClasses_Click()
    If Len(FilterGUID) > 0 Then
        FilterGUID = Empty
        LoadFile Text1
    End If
End Sub

Private Sub mnuTestInIE_Click()
    Dim cc As CClass
    On Error Resume Next
    If TypeName(ActiveNode.Tag) = "CClass" Then
        Set cc = ActiveNode.Tag
        'we need a real method name for a function on this interface to
        'trigger the safe for scripting check...
        MDIForm1.ShowWebPageForGUID cc.GUID, ActiveNode.Child.Child.Text
    End If
End Sub

Private Sub mnuViewAuditLog_Click()

    If TypeName(ActiveNode.Tag) <> "CClass" Then Exit Sub
        
    Dim cc As CClass
    Dim tmp() As String
    Dim rs As Recordset
    Dim n As String
    
    On Error Resume Next
    
    Set cc = ActiveNode.Tag
    
    If Not cc.BeenAudited Then
        MsgBox "This item has not been audited", vbInformation
    Else
        
        Set rs = QueryDsn("Select * from tblauditlog where clsid='" & cc.GUID & "'")
        
        push tmp, "Audit report for Clsid: " & cc.GUID
        push tmp, String(40, "-")
        While Not rs.EOF
            
            n = IIf(IsNull(rs!notes), "", rs!notes)
            
            push tmp, "Auditor: " & rs!auditor & vbTab & "Date: " & rs!sdate & vbTab & _
                     "Crashs: " & rs!crashs & vbTab & "Tests: " & rs!tests & vbCrLf & _
                     "Notes: " & n & vbCrLf & String(50, "-") & vbCrLf
                 
           rs.MoveNext
        Wend
        
        frmMsg.Display Join(tmp, vbCrLf)
        
    End If
    
        
        
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 And Not ActiveNode Is Nothing Then
        mnuFuzzMember.Visible = (TypeName(ActiveNode.Tag) = "CMember")
        mnuFuzzMembers.Visible = (TypeName(ActiveNode.Tag) = "CInterface")
        mnuFuzzClass.Visible = (TypeName(ActiveNode.Tag) = "CClass")
        mnuObjectSafetyReport.Visible = (TypeName(ActiveNode.Tag) = "CClass")
        mnuTestInIE.Visible = mnuFuzzClass.Visible
        mnuLoadFromMem.Visible = ((ActiveNode.Index = 1) And Len(FilterGUID) > 0)
        mnuStringScanner.Visible = (ActiveNode.Index = 1)
        PopupMenu mnuPopup
    End If
    
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim c As CMember
    Dim i As CInterface
    Dim cc As CClass
    Dim tmp() As String
    Dim report As String
    On Error Resume Next
    
    Set ActiveNode = Node
    
    mnuViewAuditLog.Visible = False
    mnuAddAuditNote.Visible = False
    
    If Node.Index = 1 Then
        push tmp(), "Loaded File: " & Text1
        push tmp(), "Name:        " & tlb.libName
        If Len(tlb.tli.GUID) > 0 Then
            push tmp(), "Lib GUID:    " & tlb.tli.GUID
            push tmp(), "Version:     " & tlb.tli.MajorVersion & "." & tlb.tli.MinorVersion
        End If
        push tmp(), "Lib Classes: " & tlb.NumClassesInLib
        Text2 = Join(tmp, vbCrLf)
    End If
    
    If TypeName(Node.Tag) = "CMember" Then
        Set c = Node.Tag
        Text2 = c.ProtoString
    End If
    
    If TypeName(Node.Tag) = "CInterface" Then
        Set i = Node.Tag
        push tmp, "Interface " & i.Name & i.DerivedString
        push tmp, "Default Interface: " & i.isDefault
        push tmp, "Members : " & i.mMembers.Count
        
        For Each c In i.mMembers
            If chkOnlyFuzzable.value = 1 Then
                If c.SupportsFuzzing Then
                    push tmp, vbTab & c.mMemberInfo.Name
                End If
            Else
                push tmp, vbTab & c.mMemberInfo.Name
            End If
        Next
        Text2 = Join(tmp, vbCrLf)
    End If
    
    If TypeName(Node.Tag) = "CClass" Then
        Set cc = Node.Tag
        push tmp, "Class " & cc.Name
        push tmp, "GUID: " & cc.GUID
        push tmp, "Number of Interfaces: " & cc.mInterfaces.Count
        push tmp, "Default Interface: " & cc.DefaultInterface
                
        If Len(cc.ObjectSafetyReport) > 0 Then
            push tmp, cc.ObjectSafetyReport
        Else
            push tmp, "RegKey Safe for Script: " & cc.SafeForScripting
            push tmp, "RegkeySafe for Init: " & cc.SafeForInitilization
        End If
                
        push tmp, "KillBitSet: " & cc.KillBitSet
        
        If Options.DistributedMode Then
            push tmp, "Audited: " & cc.BeenAudited
            mnuViewAuditLog.Visible = True
            mnuAddAuditNote.Visible = True
        End If
        
        Text2 = Join(tmp, vbCrLf)
    End If
    
End Sub

Private Sub mnuStringScanner_Click()
    Dim i As Long
    Dim tmp() As String
    Dim n As Node
    Dim m As CMember
    Dim a As CArgument
    Dim match As String
    
    On Error Resume Next
    
    match = InputBox("Enter comma delimited substrings to find", , "file,path,url,key")
    If Len(match) = 0 Then Exit Sub
    
    For Each n In tv.Nodes
        If IsObject(n.Tag) Then
            If TypeName(n.Tag) = "CMember" Then
                Set m = n.Tag
                If AnyOfTheseInstr(m.mMemberInfo.Name, match) Then
                    push tmp, "Clsid: " & m.ClassGUID & " function: " & m.mMemberInfo.Name
                End If
                For Each a In m.Args
                    If AnyOfTheseInstr(a.Name, match) Then
                        push tmp, "Clsid: " & m.ClassGUID & " function: " & m.mMemberInfo.Name & " Argument: " & a.Name
                    End If
                Next
            End If
        End If
    Next
                                     
    If Not AryIsEmpty(tmp) Then
        frmMsg.Display "Search results for match string: " & match & vbCrLf & vbCrLf & Join(tmp, vbCrLf)
    Else
        MsgBox "no string matchs found for function names or arguments :(", vbInformation
    End If
                    
End Sub

Function GetParentClass(member As Node) As CClass

    Dim rep As Long
    Dim mNode As Node
    Dim cc As CClass
    On Error Resume Next
    
    Set mNode = member
top:

    If TypeName(mNode.Tag) = "CClass" Then
        Set cc = mNode.Tag
        Set GetParentClass = cc
    Else
        rep = rep + 1
        If rep < 3 Then
            Set mNode = mNode.Parent
            GoTo top
        End If
    End If



End Function

Sub ScanElementsFor(match As String, tmp() As String, alerted As Collection)
    On Error Resume Next
    Dim key As String
    Dim n As Node
    Dim m As CMember
    Dim a As CArgument
    
    For Each n In tv.Nodes
        If IsObject(n.Tag) Then
            If TypeName(n.Tag) = "CMember" Then
                Set m = n.Tag
                If AnyOfTheseInstr(m.mMemberInfo.Name, match) Then
                    key = m.ClassGUID & "." & m.mMemberInfo.Name
                    If Not KeyExistsInCollection(alerted, key) Then
                        alerted.Add key, key
                        push tmp, "Library: " & tv.Nodes(1).Text & " - " & Text1
                        push tmp, "Class: " & GetParentClass(n).Name & "  " & m.ClassGUID & vbCrLf
                        push tmp, m.ProtoString & vbCrLf
                        push tmp, String(40, "-")
                    End If
                End If
                For Each a In m.Args
                    If AnyOfTheseInstr(a.Name, match) Then
                        key = m.ClassGUID & "." & m.mMemberInfo.Name
                        If Not KeyExistsInCollection(alerted, key) Then
                            alerted.Add key, key
                            push tmp, "Library: " & tv.Nodes(1).Text & " - " & Text1
                            push tmp, "Class: " & GetParentClass(n).Name & "  " & m.ClassGUID & vbCrLf
                            push tmp, m.ProtoString & vbCrLf
                            push tmp, String(40, "-")
                        End If
                    End If
                Next
            End If
        End If
    Next
    
End Sub

Function scerr() As String
    scerr = vbCrLf & vbCrLf & _
             "Source: " & sc.Error.Source & vbCrLf & _
            "Line: " & sc.Error.Line & " : " & sc.Error.Description & vbCrLf & _
            "Script: " & sc.Error.Text

End Function
