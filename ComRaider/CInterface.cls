VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterface"
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

Public Name As String
Public GUID As String
Public hasIDispatch As Boolean  'required because we use IDispatch::Invoke for fuzzing
Public isPublic As Boolean
Public mMembers As New Collection ' as CMember
Public DerivedString As String
Public isDefault As Boolean

Private Ignored As New Collection

Friend Sub LoadMembers(mm As Members, ownerClassGuid As String)
        Dim m As MemberInfo
        Dim cm As CMember
        Dim ignoreThreshold As Integer
        Dim i As Long
        
        If Not hasIDispatch Then
            Name = Name & " - No IDisp"
            Exit Sub
        End If
        
        
        'i = 1
        'ignoreThreshold = 3 'for IUnknown methods (must be present)
        'ignoreThreshold = ignoreThreshold + 4 'skip idisp meths too

        For Each m In mm
            'If i > ignoreThreshold Then
            If Not KeyExistsInCollection(Ignored, m.Name) Then
                dbg vbTab & "Method: " & m.Name
                Set cm = New CMember
                cm.LoadMember m, ownerClassGuid
                'Debug.Print cm.ProtoString
                mMembers.Add cm
            End If
            i = i + 1
        Next
        
End Sub


Private Sub Class_Initialize()
    
    Dim IgnoreMe() As String
    Dim x
    
    'IUnknown and IDispatch methods
    
    Const IgnoreMeths = "QueryInterface,AddRef,Release,GetTypeInfoCount," & _
                        "GetTypeInfo,GetIDsOfNames,Invoke"
    
    IgnoreMe() = Split(IgnoreMeths, ",")
    
    For Each x In IgnoreMe
        Ignored.Add x, x
    Next
        
End Sub
