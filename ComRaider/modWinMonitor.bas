Attribute VB_Name = "modWinMonitor"
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


Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef hINst As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Const WM_CLICK = &HF5

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Windows As Collection 'of CWindow


Public Function GetClass(hwnd As Long) As String
    Dim lpClassname As String, retVal As Long
    lpClassname = Space(256)
    retVal = GetClassName(hwnd, lpClassname, 256)
    GetClass = Left$(lpClassname, retVal)
End Function

Function GetCaption(hwnd As Long)
    Dim l As Long, hWndTitle As String
    l = GetWindowTextLength(hwnd)
    hWndTitle = String(l, 0)
    GetWindowText hwnd, hWndTitle, (l + 1)
    GetCaption = hWndTitle
End Function

Function ProcessFromHwnd(hwnd As Long) As String
    Dim hProc As Long, pid As Long
    Dim hMods() As Long, ret As Long, retMax As Long
    Dim sPath As String
    
    GetWindowThreadProcessId hwnd, pid
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid)
    
    If hProc <> 0 Then
        ReDim hMods(900)
        ret = EnumProcessModules(hProc, hMods(0), 900, retMax)
        sPath = Space$(260)
        ret = GetModuleFileNameExA(hProc, hMods(0), sPath, 260)
        ProcessFromHwnd = Left$(sPath, ret)
        Call CloseHandle(hProc)
    End If
    
End Function

Function EnumChildren(hwnd As Long) As Collection
    Set Windows = New Collection
    EnumChildWindows hwnd, AddressOf EnumChildProc, ByVal 0
    Set EnumChildren = Windows
End Function

Private Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim w As New CWindow
    
    w.hwnd = hwnd
    w.caption = GetCaption(hwnd)
    w.Class = GetClass(hwnd)
    w.Parent = GetParent(hwnd)
    
    Windows.Add w
    EnumChildProc = 1 'continue enum
    
End Function


