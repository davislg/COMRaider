Attribute VB_Name = "modPEParse"
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

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type IMAGE_NT_HEADERS
    Signature As String * 4
    'type IMAGE_FILE_HEADER
        Machine As Integer
        NumberOfSections As Integer
        TimeDateStamp As Long
        PointerToSymbolTable As Long
        NumberOfSymbols As Long
        SizeOfOptionalHeader As Integer
        Characteristics As Integer
    'end type
    'type IMAGE_OPTIONAL_HEADER
        Magic As Integer
        MajorLinkerVersion As Byte
        MinorLinkerVersion As Byte
        sizeOfCode As Long
        SizeOfInitializedData As Long
        SizeOfUninitializedData As Long
        AddressOfEntryPoint As Long
        BaseOfCode As Long
        BaseOfData As Long
        ImageBase As Long
        SectionAlignment As Long
        FileAlignment As Long
        MajorOperatingSystemVersion As Integer
        MinorOperatingSystemVersion As Integer
        MajorImageVersion As Integer
        MinorImageVersion As Integer
        MajorSubsystemVersion As Integer
        MinorSubsystemVersion As Integer
        Win32VersionValue As Long
        SizeOfImage As Long
        SizeOfHeaders As Long
        CheckSum As Long
        Subsystem As Integer
        DllCharacteristics As Integer
        SizeOfStackReserve As Long
        SizeOfStackCommit As Long
        SizeOfHeapReserve As Long
        SizeOfHeapCommit As Long
        LoaderFlags As Long
        NumberOfRvaAndSizes As Long
        ExportTableRVA As Long
        ExportSize As Long
    'end type
End Type


Function GetDllStats(c As CCrashMon, hFile As Long, BaseOfDll As Long, lpszPath As String, SizeOfImage As Long) As Boolean

    Const FILE_BEGIN = 0
    Dim readLen As Long
    Dim o As OVERLAPPED
    Dim dos As IMAGEDOSHEADER
    Dim nt As IMAGE_NT_HEADERS
    Dim buf() As Byte
    Dim va As Long
    Dim pName As Long
    
    Const IMAGE_NAME_RVA_OFFSET = 12
    
    ReDim buf(500)
    
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    ReadFile hFile, dos, Len(dos), readLen, o
    
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    ReadFile hFile, buf(0), 500, readLen, o
    
    CopyMemory nt, buf(dos.e_lfanew), Len(nt)
    
    va = BaseOfDll + nt.ExportTableRVA + IMAGE_NAME_RVA_OFFSET
    c.ReadLng va, pName
    
    va = BaseOfDll + pName
    c.ReadBuf va, 255, buf
    
    lpszPath = StrConv(buf, vbUnicode)
    
    If InStr(lpszPath, Chr(0)) > 0 Then
        lpszPath = Mid(lpszPath, 1, InStr(lpszPath, Chr(0)) - 1)
    End If
    
    SizeOfImage = nt.SizeOfImage
    GetDllStats = True
    
    
End Function
                                          
