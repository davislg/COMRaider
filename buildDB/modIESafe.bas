Attribute VB_Name = "modIESafe"
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
 

Private Declare Function TestClsID Lib "crashmon.dll" (ByVal clsid As Long, ByVal flags As Long, dispOk As Long, dispExOk As Long, perstOk As Long, perstStreamOk As Long, perstStorageOk As Long) As Long

Private Const CLSCTX_INPROC_SERVER As Long = 1&
Private Const CLSCTX_LOCAL_SERVER = 4
Private Const INTERFACESAFE_FOR_UNTRUSTED_CALLER = 1 'Caller of interface may be untrusted
Private Const INTERFACESAFE_FOR_UNTRUSTED_DATA = 2   'Data passed into interface may be untrusted
Private Const INTERFACE_USES_DISPEX = 4                'Object knows to use IDispatchEx")
Private Const INTERFACE_USES_SECURITY_MANAGER = 8      'Object knows to use IInternetHostSecurityManager

Private Const SafeForScriptCat = "\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}"
Private Const SafeForInitCat = "\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}"

Public reg As New clsRegistry2
Global cn As New Connection
Global fso As New CFileSystem2


Function isIOSafe(ByVal clsid As String) As CClassSafety
    
    Dim cc As New CClassSafety
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, flags As Long
     
    On Error Resume Next
    
    clsid = Trim(clsid)
    If Right(clsid, 1) <> "}" Then clsid = clsid & "}"
    If Left(clsid, 1) <> "{" Then clsid = "{" & clsid
    cc.clsid = clsid

    reg.hive = HKEY_CLASSES_ROOT
    
    cc.wasRegistered = reg.keyExists("\CLSID\" & clsid)
    If Not cc.wasRegistered Then GoTo exitNow
    
    If reg.keyExists("\CLSID\" & clsid & "\" & SafeForScriptCat) Then
        cc.regSafeForScript = True
    End If
    
    If reg.keyExists("\CLSID\" & clsid & "\" & SafeForInitCat) Then
        cc.regSafeForInit = True
    End If
    
    flags = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER
    
    If TestClsID(StrPtr(clsid), flags, a, b, c, d, e) <> 1 Then GoTo exitNow
    
    cc.HasIObjSafety = True
    cc.IDispSafe = a
    cc.IDispExSafe = b
    cc.IPersistSafe = c
    cc.IPSteamSafe = d
    cc.IPStorageSafe = e
    
exitNow:
     cc.SetGeneralFlag
     Set isIOSafe = cc
     
End Function
