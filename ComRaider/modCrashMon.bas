Attribute VB_Name = "modCrashMon"
Option Explicit
 
Global Events As New Collection
Global ProcessInfo As CREATE_PROCESS_DEBUG_INFO

Declare Function ActivePID Lib "crashmon.dll" () As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As Any) As Long
Public Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As Any) As Long
Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_QUERY_INFORMATION = (&H400)

Public Const WM_COPYDATA = &H4A
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Private Const MAXIMUM_SUPPORTED_EXTENSION = 512
Private Const SIZE_OF_80387_REGISTERS = 80

Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

'Public Type CREATE_PROCESS_DEBUG_INFO
'    hFile As Long
'    hProcess As Long
'    hThread As Long
'    lpBaseOfImage As Long
'    dwDebugInfoFileOffset As Long
'    nDebugInfoSize As Long
'    lpThreadLocalBase As Long
'    lpStartAddress As Long
'    lpImageName As Long
'    fUnicode As Integer
'End Type

Type DEBUG_EVENT
    dwDebugEventCode As DbgEvents
    dwProcessId As Long
    dwThreadId As Long
    Data(20) As Long 'enough spacen UNION***NOT SUPPORTED BY VB
End Type

'Public Type FLOATING_SAVE_AREA
'    ControlWord As Long
'    StatusWord As Long
'    TagWord As Long
'    ErrorOffset As Long
'    ErrorSelector As Long
'    DataOffset As Long
'    DataSelector As Long
'    RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte
'    Cr0NpxState As Long
'End Type
'
'Enum Registers
'    Edi
'    Esi
'    Ebx
'    Edx
'    Ecx
'    Eax
'    Ebp
'    Eip
'    Esp
'End Enum
'
'Public Type CONTEXT
'    ContextFlags As Long 'control which records returned/set
'    'CONTEXT_DEBUG_REGISTERS (NOT included in CONTEXT_FULL)
'    Dr0 As Long
'    Dr1 As Long
'    Dr2 As Long
'    Dr3 As Long
'    Dr6 As Long
'    Dr7 As Long
'    'CONTEXT_FLOATING_POINT.
'    FloatSave As FLOATING_SAVE_AREA
'    'CONTEXT_SEGMENTS.
'    SegGs As Long
'    SegFs As Long
'    SegEs As Long
'    SegDs As Long
'    'CONTEXT_INTEGER.
'    Edi As Long
'    Esi As Long
'    Ebx As Long
'    Edx As Long
'    Ecx As Long
'    Eax As Long
'    'CONTEXT_CONTROL.
'    Ebp As Long
'    Eip As Long
'    SegCs As Long       '       // MUST BE SANITIZED
'    EFlags As Long      '       // MUST BE SANITIZED 'EFlags=&H100 For Single-Step Execution!!!!!!!!!
'    Esp As Long
'    SegSs As Long
'   'CONTEXT_EXTENDED_REGISTERS.
'   ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte
'
'End Type

Public Type EXCEPTION_RECORD
   ExceptionCode As Long
   ExceptionFlags As Long
   pExceptionRecord As Long
   ExceptionAddress As Long
   NumberParameters As Long
   ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Public Type EXCEPTION_DEBUG_INFO
    ExceptionRecord As EXCEPTION_RECORD
    dwFirstChance As Long
End Type


Enum DbgEvents
         EXCEPTION_DEBUG_EVENT = 1
         EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
         EXCEPTION_SINGLE_STEP = &H80000004
         EXCEPTION_ACCESS_VIOLATION = &HC0000005
         EXCEPTION_BREAKPOINT = &H80000003
         EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
         EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
         EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
         EXCEPTION_FLT_OVERFLOW = &HC0000091
         EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
         EXCEPTION_INT_OVERFLOW = &HC0000095
         EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
         EXCEPTION_PRIV_INSTRUCTION = &HC0000096
         CREATE_THREAD_DEBUG_EVENT = 2
         CREATE_PROCESS_DEBUG_EVENT = 3
         EXIT_THREAD_DEBUG_EVENT = 4
         EXIT_PROCESS_DEBUG_EVENT = 5
         LOAD_DLL_DEBUG_EVENT = 6
         UNLOAD_DLL_DEBUG_EVENT = 7
         OUTPUT_DEBUG_STRING_EVENT = 8
End Enum

'Enum ctx_vals
'     CONTEXT_i486 = &H10000
'     CONTEXT_CONTROL = 1
'     CONTEXT_INTEGER = 2
'     CONTEXT_SEGMENTS = 4
'     CONTEXT_FLOATING_POINT = 8
'     CONTEXT_DEBUG_REGISTERS = 16
'     CONTEXT_EXTENDED_REGISTERS = 32
'     CONTEXT_EXTENDED_INTEGER = (CONTEXT_INTEGER Or &H10)
'     CONTEXT_FULL = (CONTEXT_CONTROL Or CONTEXT_FLOATING_POINT Or CONTEXT_INTEGER Or CONTEXT_EXTENDED_INTEGER)
'End Enum


Function GetEvent(id As Long) As String
    On Error GoTo hell
    If Events.Count = 0 Then LoadEvents
    GetEvent = Events("id:" & id)
    Exit Function
hell: GetEvent = "Unknown id: " & id
End Function



Sub LoadEvents()
    
    Events.Add "DEBUG_EVENT", "id:" & EXCEPTION_DEBUG_EVENT
    Events.Add "DATATYPE_MISALIGNMENT", "id:" & EXCEPTION_DATATYPE_MISALIGNMENT
    Events.Add "SINGLE_STEP", "id:" & EXCEPTION_SINGLE_STEP
    Events.Add "ACCESS_VIOLATION", "id:" & EXCEPTION_ACCESS_VIOLATION
    Events.Add "BREAKPOINT", "id:" & EXCEPTION_BREAKPOINT
    Events.Add "ARRAY_BOUNDS_EXCEEDED", "id:" & EXCEPTION_ARRAY_BOUNDS_EXCEEDED
    Events.Add "FLT_DIVIDE_BY_ZERO", "id:" & EXCEPTION_FLT_DIVIDE_BY_ZERO
    Events.Add "FLT_INVALID_OPERATION", "id:" & EXCEPTION_FLT_INVALID_OPERATION
    Events.Add "FLT_OVERFLOW", "id:" & EXCEPTION_FLT_OVERFLOW
    Events.Add "INT_DIVIDE_BY_ZERO", "id:" & EXCEPTION_INT_DIVIDE_BY_ZERO
    Events.Add "ILLEGAL_INSTRUCTION", "id:" & EXCEPTION_ILLEGAL_INSTRUCTION
    Events.Add "PRIV_INSTRUCTION", "id:" & EXCEPTION_PRIV_INSTRUCTION
    Events.Add "CREATE_THREAD", "id:" & CREATE_THREAD_DEBUG_EVENT
    Events.Add "CREATE_PROCESS", "id:" & CREATE_PROCESS_DEBUG_EVENT
    Events.Add "EXIT_THREAD", "id:" & EXIT_THREAD_DEBUG_EVENT
    Events.Add "EXIT_PROCESS", "id:" & EXIT_PROCESS_DEBUG_EVENT
    Events.Add "LOAD_DLL", "id:" & LOAD_DLL_DEBUG_EVENT
    Events.Add "UNLOAD_DLL", "id:" & UNLOAD_DLL_DEBUG_EVENT
    Events.Add "DEBUG_STRING", "id:" & OUTPUT_DEBUG_STRING_EVENT
    
    
End Sub



'Public Function GetContext(ByVal threadID As Long) As CONTEXT
'    Dim ThHandle As Long
'    ThHandle = tm.GetThreadHandle(threadID)
'    GetContext.ContextFlags = CONTEXT_i486 Or CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT
'    GetThreadContext ThHandle, GetContext
'End Function
'
'Public Sub SetContext(ByVal threadID As Long, NewContext As CONTEXT)
'    Dim ThHandle As Long
'    ThHandle = tm.GetThreadHandle(threadID)
'    SetThreadContext ThHandle, NewContext
'End Sub
 



Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function
'
'
'
'Function shexdump(it)
'    Dim my, i, c, s, a, b
'    Dim lines() As String
'
'    my = ""
'    For i = 1 To Len(it)
'        a = Asc(Mid(it, i, 1))
'        c = Hex(a)
'        c = IIf(Len(c) = 1, "0" & c, c)
'        b = b & IIf(a > 65 And a < 120, Chr(a), ".")
'        my = my & c & " "
'        If i Mod 16 = 0 Then
'            push lines(), my & "  [" & b & "]"
'            my = Empty
'            b = Empty
'        End If
'    Next
'
'    If Len(b) > 0 Then
'        If Len(my) < 48 Then
'            my = my & String(48 - Len(my), " ")
'        End If
'        If Len(b) < 16 Then
'             b = b & String(16 - Len(b), " ")
'        End If
'        push lines(), my & "  [" & b & "]"
'    End If
'
'    If Len(it) < 16 Then
'        shexdump = my & "  [" & b & "]" & vbCrLf
'    Else
'        shexdump = Join(lines, vbCrLf)
'    End If
'
'
'End Function
'
'Function hexdump(ByVal base As Long, it() As Byte)
'    Dim my, i, c, s, a As Byte, b
'    Dim lines() As String
'
'    my = ""
'    For i = 0 To UBound(it)
'        a = it(i)
'        c = Hex(a)
'        c = IIf(Len(c) = 1, "0" & c, c)
'        b = b & IIf(a > 65 And a < 120, Chr(a), ".")
'        my = my & c & " "
'        If (i + 1) Mod 16 = 0 Then
'            push lines(), Hex(base) & " " & my & " [" & b & "]"
'            base = base + 16
'            my = Empty
'            b = Empty
'        End If
'    Next
'
'    If Len(b) > 0 Then
'        If Len(my) < 48 Then
'            my = my & String(48 - Len(my), " ")
'        End If
'        If Len(b) < 16 Then
'             b = b & String(16 - Len(b), " ")
'        End If
'        push lines(), my & " [" & b & "]"
'    End If
'
'    If UBound(it) < 16 Then
'        hexdump = Hex(base) & " " & my & " [" & b & "]" & vbCrLf
'    Else
'        hexdump = Join(lines, vbCrLf)
'    End If
'
'
'End Function
'
'
'
'Function FileExists(path) As Boolean
'  If Len(path) = 0 Then Exit Function
'  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
'  Else FileExists = False
'End Function
'
'Function WriteByte(va As Long, b As Byte) As Long
'    WriteByte = WriteProcessMemory(ProcessInfo.hProcess, ByVal va, b, 1, 0&)
'End Function
'
''untested
'Function WriteBuf(va As Long, b() As Byte) As Long
'    WriteBuf = WriteProcessMemory(ProcessInfo.hProcess, ByVal va, b(0), UBound(b) + 1, 0&)
'End Function
'
'Function WriteLng(va As Long, v As Long) As Long
'    WriteLng = WriteProcessMemory(ProcessInfo.hProcess, ByVal va, v, 4, 0&)
'End Function
'
'Function ReadByte(va As Long) As Byte
'    ReadProcessMemory ProcessInfo.hProcess, va, ReadByte, 1, 0
'End Function
'
'Function ReadLng(va As Long) As Long
'    Dim b(4) As Byte
'    ReadProcessMemory ProcessInfo.hProcess, va, b(0), 4, 0
'    CopyMemory ReadLng, b(0), 4
'End Function
'
'Function ReadBuf(va As Long, leng As Long) As Byte()
'    Dim tmp() As Byte
'    ReDim tmp(leng - 1)
'    ReadProcessMemory ProcessInfo.hProcess, va, tmp(0), leng, 0
'    ReadBuf = tmp
'End Function
 

 


