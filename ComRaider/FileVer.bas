Attribute VB_Name = "FileProps"
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
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

'Used in several projects do not change interface!

Option Explicit

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Function FileVersion(ByVal sPath As String) As String
    Dim tmp
    On Error Resume Next
    tmp = Split(QuickInfo(sPath), vbCrLf)
    tmp = Split(tmp(2), vbTab)
    FileVersion = tmp(UBound(tmp))
End Function

Function QuickInfo(ByVal PathWithFilename As String)
    'return file-properties of given file  (EXE , DLL , OCX)
    'adapted from: http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then
        Exit Function
    End If
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim bytBuffer() As Byte
    Dim bytBuff() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strVersionInfo() As String
    Dim strTemp As String
    Dim intTemp As Integer
    Dim n As Long
           
    On Error Resume Next
    
    ReDim bytBuff(500)
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    
    If lngBufferlen > 0 Then
    
       ReDim bytBuffer(lngBufferlen)
       
       If GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0)) <> 0 Then
          If VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen) <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)

             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             
             Const fields = "CompanyName,FileDescription,FileVersion,InternalName,LegalCopyright,OriginalFileName,ProductName,ProductVersion"
             strVersionInfo = Split(fields, ",")
             
             For intTemp = 0 To UBound(strVersionInfo)
                strBuffer = String$(800, 0)
                strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(intTemp)
                If VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen) = 0 Then
                   strVersionInfo(intTemp) = ""
                Else
                   lstrcpy strBuffer, lngVerPointer
                   strVersionInfo(intTemp) = strVersionInfo(intTemp) & vbTab & vbTab & Replace(strBuffer, Chr(0), Empty)
                End If
             Next
             
          End If
       End If
    End If
    
    QuickInfo = Join(strVersionInfo, vbCrLf)
    
End Function
 
