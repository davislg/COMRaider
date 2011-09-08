VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClass"
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

'Private Type strucGUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(7) As Byte
'End Type
'
'Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As strucGUID) As Long
'Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As strucGUID) As Long
'Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (pCLSID As strucGUID, lpszProgID As Long) As Long
'Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As strucGUID, lpszProgID As Long) As Long
'Private Declare Sub CoCreateInstance Lib "ole32.dll" (ByRef rclsid As strucGUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As strucGUID, ppv As Any)
'
'Private IID_IUnknown As strucGUID
'Private Const IIDSTR_IUnknown = "{00000000-0000-0000-C000-000000000046}"
'Private Const CLSCTX_INPROC_SERVER As Long = 1&

Public Name As String
Public GUID As String
Public ProgID As String
Public BeenAudited As Boolean
Public mInterfaces  As New Collection 'as CInterface
Public SafeForScripting As Boolean
Public SafeForInitilization As Boolean
Public KillBitSet As Boolean
Public DefaultInterface As String
Public ObjectSafetyReport As String

'Function CreateInstance() As IUnknown
'
'    Dim udtGUID As strucGUID
'    Dim iUnk As IUnknown
'
'    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
'    CLSIDFromString StrPtr(GUID), udtGUID
'    CoCreateInstance udtGUID, 0, CLSCTX_INPROC_SERVER, IID_IUnknown, iUnk
'
'    Set CreateInstance = iUnk
'
'End Function


Friend Sub LoadInterfaces(ii As Interfaces, defInterface As String)
        
        Dim iface As InterfaceInfo
        Dim mi As CInterface
        Dim cnt As Long
          
        reg.hive = HKEY_CLASSES_ROOT
'        If reg.keyExists(clsid & GUID & "\Programmable") Then Programmable = True

        If Not GetSavedSafetyReport(GUID, ObjectSafetyReport) Then
            Const tail = "\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}"
            Const tail2 = "\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}"
            If reg.keyExists("\CLSID\" & GUID & tail) Then SafeForScripting = True
            If reg.keyExists("\CLSID\" & GUID & tail2) Then SafeForInitilization = True
        End If
              
        If KeyExistsInCollection(killbitted, GUID) Then KillBitSet = True
        
        DefaultInterface = defInterface
        
        For Each iface In ii
            AddInterface iface
        Next
        
End Sub

Sub AddInterface(iface As InterfaceInfo)
        
        On Error Resume Next
        Dim mi As New CInterface

        mi.DerivedString = GetDerivedInterfaces(iface, mi)
        dbg "Interface: " & iface.Name & " " & mi.DerivedString & IIf(mi.isPublic, " : Public", " : Private")
        
        If Not mi.isPublic Then Exit Sub
        
        With iface
            mi.GUID = .GUID
            mi.Name = .Name
            If .Name = DefaultInterface Then mi.isDefault = True
            mi.LoadMembers .Members, Me.GUID
            mInterfaces.Add mi
            
            If Options.DistributedMode Then
                cnt = QueryDsn("Select count(autoid) as cnt from tblauditlog where clsid='" & Me.GUID & "'")!cnt
                If cnt > 0 Then Me.BeenAudited = True
            End If
        End With
        
End Sub


'Private Function isControl(iface As InterfaceInfo) As Boolean
'    Dim flags As Long
'    If (iface.AttributeMask And TYPEFLAG_FCONTROL) = 0 Then isControl = True
'End Function

Private Function isPublic(iface As InterfaceInfo) As Boolean
    Dim flags As Long
    flags = IMPLTYPEFLAG_FSOURCE Or IMPLTYPEFLAG_FRESTRICTED
    If (iface.AttributeMask And flags) = 0 Then isPublic = True
End Function

Private Function GetDerivedInterfaces(iface As InterfaceInfo, mi As CInterface) As String
    Dim ii As InterfaceInfo
    Dim tmp As String
    
    If isPublic(iface) Then mi.isPublic = True
    
    For Each ii In iface.ImpliedInterfaces
        If ii.Name = "IDispatch" Then mi.hasIDispatch = True
        tmp = tmp & " : " & ii.Name
    Next
    
    GetDerivedInterfaces = tmp
    
End Function
