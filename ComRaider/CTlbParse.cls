VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTlbParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Declare Function CreateIDispObj Lib "crashmon.dll" (ByVal clsid As Long, instance As Object) As Boolean
Private Declare Function GetInterfaceInfoForClsid Lib "crashmon.dll" (ByVal clsid As Long, instance As Object, tmpObj As Object) As Boolean

Public tli As New TypeLibInfo
Public mClasses As New Collection
Public ErrMsg As String
Public LibName As String
Public LoadedFromLiveInst As Boolean
Public NO_VALID_TLB As Boolean
Public NumClassesInLib As Long


Function LoadFile(path As String, Optional targetGUID As String) As Boolean
    Dim ccinfo As CoClassInfo
    Dim c As CClass

    On Error Resume Next
    tli.ContainingFile = path
    
    If Err.Number <> 0 Then
        ErrMsg = Err.Description
        If InStr(Err.Description, "file doesn't contain a valid typelib") > 0 Then
            ErrMsg = "File does not contain a valid typelib"
            NO_VALID_TLB = True
        End If
        Exit Function
    End If
    
    LibName = tli.Name
    NumClassesInLib = tli.CoClasses.Count
    
    For Each ccinfo In tli.CoClasses
    
        If Len(targetGUID) > 0 Then
            If LCase(ccinfo.GUID) <> LCase(targetGUID) Then GoTo nextOne
        End If
        
        Set c = New CClass
        With ccinfo
            c.Name = .Name
            c.GUID = .GUID
            c.LoadInterfaces ccinfo.Interfaces, ccinfo.DefaultInterface
            mClasses.Add c
        End With
        
nextOne:
    Next
    LoadFile = True

End Function

Function LoadFromMem(targetGUID As String) As Boolean
    
        On Error Resume Next
        Dim ii As InterfaceInfo
        Dim c As CClass
        Dim o As Object
        Dim ccinfo As CoClassInfo
        
        'If Not CreateIDispObj(StrPtr(targetGUID), o) Then
        '    ErrMsg = "Could not create object clsid: " & targetGUID & vbCrLf & vbCrLf & "It probably does not support IDispatch"
        '    Exit Function
        'End If
                
        'Set ii = BestInterfaceInfo(o)
        
        GetInterfaceInfoForClsid StrPtr(targetGUID), ii, o
             
        If o Is Nothing Then
             ErrMsg = "Could not create object clsid: " & targetGUID & vbCrLf & vbCrLf & _
                      "It probably does not support IDispatch"
             Exit Function
        End If
        
        If ii Is Nothing Then
             ErrMsg = "TLIApplication.InterfaceInfoFromObject failed"
             Exit Function
        End If
        
        Set c = New CClass
        NumClassesInLib = ii.Parent.CoClasses.Count
        LibName = ii.Parent.Name
        
        If ii.Parent.CoClasses.Count = 0 Then
            With ii
                c.Name = .Name
                c.GUID = targetGUID
                c.AddInterface ii
                c.DefaultInterface = ii.Name
                mClasses.Add c
            End With
            Debug.Print "No other CoClasses found? Typename(o)= " & TypeName(o) & " libname: " & LibName
        Else
            For Each ccinfo In ii.Parent.CoClasses
                If LCase(ccinfo.GUID) <> LCase(targetGUID) Then GoTo nextOne
                Set c = New CClass
                With ccinfo
                    c.Name = .Name
                    c.GUID = .GUID
                    c.LoadInterfaces ccinfo.Interfaces, ccinfo.DefaultInterface
                    mClasses.Add c
                End With
nextOne:
            Next
        End If
        
    
        LoadFromMem = True

End Function

 
'this should be done in vc dll because it can make the whole app barf on occasion!
'Private Function BestInterfaceInfo(ByVal obj As Object) As InterfaceInfo
'  Dim ta As New TLIApplication
'
'  On Error GoTo NotAvailable
'
'  Set BestInterfaceInfo = ta.InterfaceInfoFromObject(obj)
'
'  'With BestInterfaceInfo.Parent 'TypeLibInfo object
'  '      With ta.TypeLibInfoFromRegistry(.GUID, .MajorVersion, .MinorVersion, .LCID)
'  '          Set BestInterfaceInfo = .Me.TypeInfos.IndexedItem(BestInterfaceInfo.TypeInfoNumber)
'  '      End With
'  'End With
'
'  Exit Function
'NotAvailable: ErrMsg = Err.Description
'End Function



