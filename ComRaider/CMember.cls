VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMember"
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

Option Explicit

'this class encapsulates one function as defined in the typelib
'this class is owned by an CInterface class and held as part of its mMembers col
'this class in turn holds a direct ref to the member tli obj as well as a
'simplified and parsed collection of CArgument classes used by .Invoke to help
'call it


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Enum FType
    ftsub = 0
    ftFunc = 1
    ftPropLet = 2
    ftPropGet = 3
    ftPropSet = 4
    ftUnknown = 5
End Enum

Public ClassGUID As String
Public CallType As Long
Public ProtoString As String
Public mMemberInfo As MemberInfo
Public ArgCount As Integer
Public SupportsFuzzing As Boolean 'entire method cant be fuzzed
Public Args As New Collection 'of CArgument

Private Declare Function GetTickCount Lib "kernel32" () As Long


Sub LoadMember(m As MemberInfo, ownerClassGuid As String)
    
    'If m.Name = "insertBefore" Then Stop
    
    Set mMemberInfo = m
    ClassGUID = ownerClassGuid
    ArgCount = m.Parameters.Count
    ProtoString = GetProtoString()
    
    Dim a As CArgument
    Dim X As Long
    
    For Each a In Args
        If Not a.isWorthFuzzing Then X = X + 1
    Next
    
    If X = Args.Count Then SupportsFuzzing = False
    
End Sub


Function BuildFuzzConfigFiles(log() As String, ByVal pFolder As String, Optional ByVal proLogue As String, Optional libName As String, Optional className As String) As String()      'returns filepaths to config files
    Dim fPaths() As String
    Dim i As Long
    Dim default
        
    pFolder = pFolder & "\" & mMemberInfo.Name
    If fso.FolderExists(pFolder) Then fso.DeleteFolder (pFolder) 'clean out old member scripts
    If Not fso.FolderExists(pFolder) Then fso.buildPath pFolder
    
    Dim arg As CArgument
    Dim ary()
     
    ReDim ary(1 To Args.Count)
    
    'first fill out default args for not fuzzable friends
    For i = 1 To Args.Count
        Set arg = Args(i)
        
        If Not arg.isWorthFuzzing Then
            Select Case arg.DataType 'the ones we dont bother to fuzz
                Case atByte:    ary(i) = 1
                Case atDate:    ary(i) = "Now"
                Case atBool:    ary(i) = "True"
                Case atempty:   ary(i) = """"""
                Case atSingle:  ary(i) = 1
                Case atUnknown: ary(i) = "Empty"
                Case atObject:
                                'If Len(arg.ObjectType) > 0 Then
                                '    If InStr(arg.ObjectType, ".") < 1 Then
                                '        arg.ObjectType = libName & "." & arg.ObjectType
                                '    End If
                                '    ary(i) = "CreateObject(""" & arg.ObjectType & """)"
                                'Else
                                    ary(i) = "CreateObject(""ADODB.Record"") 'just some live obj"
                                'End If
            End Select
        Else
            Select Case arg.DataType 'defaults for those we do fuzz
                Case atLong, atInteger, atdouble:    ary(i) = 1
                Case atVariant, atstring:            ary(i) = """defaultV"""
            End Select
        End If
    Next
    
    Dim X

    'now: base ary of defaults are filled out,
    '     fuzz variables loaded from app.path & tests.vbs
    '     now we build all of the fuzz files for each method and arg type
    For i = 1 To Args.Count
        Set arg = Args(i)
        If arg.isWorthFuzzing Then
            default = ary(i) 'save default val for next loop
            Select Case arg.DataType
                    Case atLong:
                            For Each X In frmtlbViewer.lngs
                                ary(i) = X
                                WSFFile pFolder, ary, fPaths, , proLogue, libName, className
                            Next
                    Case atInteger
                            For Each X In frmtlbViewer.ints
                                ary(i) = X
                                WSFFile pFolder, ary, fPaths, , proLogue, libName, className
                            Next
                    Case atdouble
                            For Each X In frmtlbViewer.dbls
                                ary(i) = X
                                WSFFile pFolder, ary, fPaths, , proLogue, libName, className
                            Next
                    Case atstring, atVariant
                            For Each X In frmtlbViewer.strs
                                ary(i) = X
                                WSFFile pFolder, ary, fPaths, , proLogue, libName, className
                            Next
                            
                            'load any custom fuzz strings for this arg type (only for variants and strings!)
                            LoadCustomFuzzArgsForVar LCase(arg.Name)
                            For Each X In frmtlbViewer.cust
                                ary(i) = X
                                WSFFile pFolder, ary, fPaths, , proLogue, libName, className
                            Next
            
            End Select
            ary(i) = default 'back to default so we can move to next
        End If
    Next

    BuildFuzzConfigFiles = fPaths 'return filepaths
        
End Function


Private Sub LoadCustomFuzzArgsForVar(argName As String)
    On Error Resume Next
    Set frmtlbViewer.cust = New Collection
    frmtlbViewer.sc.Eval "CustomStrArg(""" & argName & """)"
End Sub


Private Function WSFFile(ByVal pFolder As String, ary(), fPath() As String, Optional baseFileName As String, Optional proLogue As String, Optional libName As String, Optional className As String)
    
    Dim f As String
    Dim tmp() As String
    Dim argStr As String
    Dim ik As InvokeKinds
    Dim i As Long
    Dim mn As String
    
    '<?XML version="1.0" standalone="yes" ?>
    '<package><job id="DoneInVBS"><?job debug="true"?>
    '<object id="target" classid="clsid:D7A7D7C3-D47F-11D0-89D3-00A0C90833E6"/>
    '<script language="VBScript">
    'WScript.Echo "This is VBScript"
    'WScript.Echo TypeName(Target)
    '</script></job></package>

    If Right(pFolder, 1) <> "\" Then pFolder = pFolder & "\"
    
    If Len(baseFileName) > 0 Then
        f = pFolder & baseFileName & ".wsf"
        If fso.FileExists(f) Then f = SafeFreeFileName(pFolder, ".wsf")
    Else
        f = SafeFreeFileName(pFolder, ".wsf")
    End If
    
    push tmp, "<?XML version='1.0' standalone='yes' ?>"
    push tmp, "<package><job id='DoneInVBS' debug='false' error='true'>"
    push tmp, "<object classid='clsid:" & Mid(ClassGUID, 2, Len(ClassGUID) - 2) & "' id='target' />"
    push tmp(), "<script language='vbscript'>"
    push tmp(), ""
    push tmp(), "'File Generated by COMRaider v" & App.Major & "." & App.Minor & "." & App.Revision & " - http://labs.idefense.com"
    push tmp(), ""
    push tmp(), "'Wscript.echo typename(target)"
    push tmp, ""
    push tmp(), "'for debugging/custom prolog"
    push tmp(), "targetFile = """ & frmtlbViewer.Text1 & """" '>> assumes here only by frmtlbviewer
    push tmp(), "prototype  = """ & Replace(Replace(Replace(ProtoString, vbCrLf, Empty), vbTab, " "), "  ", " ") & """"
    push tmp(), "memberName = """ & Me.mMemberInfo.Name & """"
    push tmp(), "progid     = """ & libName & "." & className & """"
    push tmp(), "argCount   = " & Args.Count
    push tmp(), ""
    
    For i = 1 To UBound(ary)
        If Len(ary(i)) = 0 Then 'this was if ary(i) = empty but messed up on ary(i) = 0 too
            argStr = argStr & " , " 'optional parameter
        Else
            argStr = argStr & "arg" & i & " ,"
            If IsObject(ary(i)) Then
                 push tmp(), "set arg" & i & "=" & ary(i)
            Else
                 push tmp(), "arg" & i & "=" & ary(i)
            End If
        End If
    Next
    
    push tmp, ""
    If Len(proLogue) > 0 Then push tmp(), proLogue & vbCrLf
    
    mn = mMemberInfo.Name
    If Left(mn, 1) = "_" Then mn = "[" & mn & "]"
    
    If CallType = ftPropLet Then
        If UBound(ary) > 1 Then
            'Property let CellHeight(row as long, column as long) as string -> target.Cellheight(arg1,arg2) = arg3
            Dim ubarg As String
            ubarg = "arg" & UBound(ary)
            argStr = Replace(argStr, "," & ubarg & " ,", Empty)
            push tmp, "target." & mn & "(" & argStr & ") = " & ubarg
        Else
            push tmp(), "target." & mn & " = arg1"
        End If
    Else
        push tmp(), "target." & mn & " " & VBA.Left(argStr, Len(argStr) - 1)
    End If
    
    push tmp(), vbCrLf & "</script></job></package>"
    
    fso.WriteFile f, Join(tmp, vbCrLf)
    
    push fPath(), f
    
End Function







Private Function GetProtoString() As String
    
Dim pi As ParameterInfo
Dim TIResolved As TypeInfo
Dim arg As CArgument
Dim tmp() As String
Dim isFirst As Boolean
       
    SupportsFuzzing = True
    isFirst = True
     
    BuildFuncName tmp 'sets some props and builds prototype string head
       
    With mMemberInfo.Parameters
        If .Count Then
              glue tmp, "(" & vbCrLf
              
              For Each pi In mMemberInfo.Parameters
              
                    Set arg = New CArgument
                    If isFirst Then isFirst = False Else glue tmp, ", " & vbCrLf 'not first parameter
                    
                    arg.isDefault = pi.default
                    arg.isOptional = pi.Optional
                    
                    If pi.default Or pi.Optional Then
                          If .OptionalCount = -1 Then  'paramarray
                              arg.DataType = atParamAry
                              push tmp, "[ParamArray " 'This will be the only optional parameter
                              SupportsFuzzing = False
                          End If
                    End If
                    
                    With pi.VarTypeInfo
                          If (.VarType And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                              ParseObjArg tmp, pi, arg, TIResolved
                          Else
                              ParseStdArg tmp, pi, arg
                          End If
                      
                          If pi.Optional Then
                              If pi.default Then
                                  glue tmp, ProduceDefaultValue(pi.DefaultValue, TIResolved)
                              End If
                              glue tmp, "]"
                          End If
                    End With
                    
                    arg.Name = pi.Name 'added late in game!
                    Args.Add arg
              Next
              
              glue tmp, vbCrLf & ")"
        End If
    End With

    GetReturnType tmp, False

    GetProtoString = Join(tmp, vbCrLf)
  
End Function

Private Sub ParseStdArg(tmp, pi As ParameterInfo, arg As CArgument)

    Dim strTypeName As String

    With pi.VarTypeInfo
    
        arg.isByVal = IIf(.PointerLevel = 0, True, False)
        glue tmp, vbTab & IIf(arg.isOptional, "[ ", "") & IIf(arg.isByVal, "ByVal ", "ByRef ") & pi.Name
                                    
        If .VarType <> vbVariant Then
              strTypeName = TypeName(.TypedVariant)
              If .VarType And (VT_ARRAY Or VT_VECTOR) Then
                    glue tmp, "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                    If Not pi.Optional Then SupportsFuzzing = False
                    arg.ManSetType atarray, False, .VarType
              Else
                    If Not Options.AllowObjArgs And strTypeName = "Object" Then SupportsFuzzing = False
                    arg.SetStdType strTypeName
                    glue tmp, " As " & strTypeName
              End If
        Else
              arg.ManSetType atVariant, True, VT_VARIANT
              glue tmp, " As Variant"
        End If
        
    End With
                                
End Sub

Private Sub ParseObjArg(tmp, pi As ParameterInfo, arg As CArgument, TIResolved As TypeInfo)
    Dim TIType As TypeInfo
    Dim TKind As TypeKinds
 
    On Error Resume Next
    
    TKind = TKIND_MAX
    
    arg.ManSetType atObject, False ' True 'default guess might change
    
    With pi.VarTypeInfo
                   
            Set TIType = .TypeInfo
            
            If Not TIType Is Nothing Then
                
                Set TIResolved = TIType
                TKind = TIResolved.TypeKind
                
                Do While TKind = TKIND_ALIAS
                    TKind = TKIND_MAX
                    Set TIResolved = TIResolved.ResolvedType
                    If Err Then Err.Clear Else TKind = TIResolved.TypeKind
                Loop
                
            End If
              
            Select Case TKind
                  Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                          arg.isByVal = .PointerLevel = 1
                  Case TKIND_RECORD 'Records not passed ByVal in VB
                          arg.isByVal = False
                          arg.ManSetType atLong, True
                  Case TKIND_ENUM
                          arg.ManSetType atLong, True
                          arg.isByVal = .PointerLevel = 0
                  Case Else
                          arg.isByVal = .PointerLevel = 0
            End Select
            
            glue tmp, vbTab & IIf(arg.isOptional, "[ ", "") & IIf(arg.isByVal, "ByVal ", "ByRef ") & pi.Name
            
            If .VarType And (VT_ARRAY Or VT_VECTOR) Then
                glue tmp, "()"
                arg.ManSetType atarray, False
            End If
            
            If TIType Is Nothing Then 'Error
                  glue tmp, "As ?"
            Else
                  If .IsExternalType Then
                          glue tmp, "As " & .TypeLibInfoExternal.Name & "." & TIType.Name
                          arg.ObjectType = .TypeLibInfoExternal.Name & "." & TIType.Name
                  Else
                          glue tmp, "As " & TIType.Name
                          arg.ObjectType = TIType.Name
                  End If
            End If

    End With
    
    If arg.DataType = atObject And Not Options.AllowObjArgs Then
        If Not pi.Optional Then SupportsFuzzing = False
    End If
                                
End Sub
Private Sub BuildFuncName(tmp)
    
    With mMemberInfo
        If .InvokeKind = INVOKE_FUNC Or .InvokeKind = INVOKE_EVENTFUNC Then
            Select Case .ReturnType.VarType
                  Case VT_VOID, VT_HRESULT
                      push tmp, "Sub"
                      CallType = ftsub
                      If .Parameters.Count = 0 Then SupportsFuzzing = False
                  Case Else
                      push tmp, "Function"
                      CallType = ftFunc
                      If .Parameters.Count = 0 Then SupportsFuzzing = False
            End Select
        Else
        
            Select Case .InvokeKind
                Case INVOKE_PROPERTYGET
                        CallType = ftPropGet
                        SupportsFuzzing = False
                        push tmp, "Property Get"
                        
                Case INVOKE_PROPERTYPUT
                        CallType = ftPropLet
                        SupportsFuzzing = True
                        push tmp, "Property Let"
                        
                Case INVOKE_PROPERTYPUTREF
                        CallType = ftPropSet
                        SupportsFuzzing = False
                        push tmp, "Property Set"
                
                Case INVOKE_UNKNOWN
                        CallType = ftPropLet 'experimental treating these as PROPPUT
                        SupportsFuzzing = True
                        'CallType = ftUnknown
                        'SupportsFuzzing = False
                        push tmp, "Invoke_Unknown "
                        
            End Select
            
        End If
        
        glue tmp, .Name
    End With
    
End Sub

Private Function GetReturnType(tmp, fIsConstant As Boolean)
    Dim strTypeName As String
    Dim VarTypeCur As Integer
    Dim ConstVal
    Dim arg As CArgument
   
    With mMemberInfo
    
        If fIsConstant Then
               ConstVal = .value
               glue tmp, " = " & ConstVal
               Select Case VarType(ConstVal)
                 Case vbInteger, vbLong
                   If ConstVal < 0 Or ConstVal > 15 Then
                       glue tmp, " (&H" & Hex$(ConstVal) & ")"
                   End If
               End Select
         Else
               With .ReturnType
                 VarTypeCur = .VarType
                 If VarTypeCur = 0 Or (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                       
                       On Error Resume Next
                       
                       'we cant fuzz prop puts who want any of these
                       If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then SupportsFuzzing = False
                       
                       If Not .TypeInfo Is Nothing Then
                           
                           If Err Then 'Information not available
                               glue tmp, " As ?"
                           Else
                               If .IsExternalType Then
                                       glue tmp, " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                               Else
                                       glue tmp, " As " & .TypeInfo.Name
                               End If
                           End If
                       End If
                       
                       If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then glue tmp, "()"
                       On Error GoTo 0
                 Else
                       Select Case VarTypeCur
                         Case VT_VARIANT, VT_VOID, VT_HRESULT
                         Case Else
                           strTypeName = TypeName(.TypedVariant)
                           If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                               If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then
                                    SupportsFuzzing = False
                               End If
                               glue tmp, "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                           Else
                               glue tmp, " As " & strTypeName
                               
                               'If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then
                               If CallType = ftPropLet Then
                                    Set arg = New CArgument      'only chance to add it special case logic
                                    arg.isByVal = True
                                    arg.SetStdType strTypeName   'will handle this special case
                                    arg.isDefault = False
                                    arg.isOptional = False
                                    arg.Name = mMemberInfo.Name
                                    Args.Add arg
                                    ArgCount = 1
                               End If
                                
                           End If
                       End Select
                 End If
               End With
         End If
    End With

End Function

Private Function ProduceDefaultValue(DefVal As Variant, ByVal ti As TypeInfo) As String
Dim lTrackVal As Long
Dim mi As MemberInfo
Dim TKind As TypeKinds
    If ti Is Nothing Then
        Select Case VarType(DefVal)
            Case vbString
                If Len(DefVal) Then
                    ProduceDefaultValue = """" & DefVal & """"
                End If
            Case vbBoolean 'Always show for Boolean
                ProduceDefaultValue = DefVal
            Case vbDate
                If DefVal Then
                    ProduceDefaultValue = "#" & DefVal & "#"
                End If
            Case Else 'Numeric Values
                If DefVal <> 0 Then
                    ProduceDefaultValue = DefVal
                End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = ti.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
            Set ti = ti.ResolvedType
            If Err = 0 Then TKind = ti.TypeKind
            On Error GoTo 0
        Loop
        If ti.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each mi In ti.Members
                If mi.value = lTrackVal Then
                    ProduceDefaultValue = mi.Name
                    Exit For
                End If
            Next
        End If
    End If
End Function


'Function BuildFuzzConfigFiles(log() As String, ByVal pFolder As String) As String()   'returns filepaths to config files
'    Dim fz As fzConfig
'    Dim fPaths() As String
'    Dim i As Long
'    Dim default
'
'    With fz
'        .argCnt = ArgCount
'        .objGuid = ClassGUID
'        .InvokeKind = mMemberInfo.InvokeKind
'        .methodID = mMemberInfo.MemberId
'        .methodName = mMemberInfo.Name
'        .protoType = ProtoString
'    End With
'
'    pFolder = pFolder & "\" & mMemberInfo.Name
'    If fso.FolderExists(pFolder) Then fso.DeleteFolder (pFolder) 'clean out old member scripts
'    If Not fso.FolderExists(pFolder) Then fso.buildPath pFolder
'
'    Dim arg As CArgument
'    Dim ary()
'
'    ReDim ary(1 To ArgCount)
'
'    'first let fill out default args for not fuzzable friends
'    For i = 1 To Args.Count
'        Set arg = Args(i)
'
'        If arg.isOptional Then
'            ary(i) = Empty
'        ElseIf Not arg.isWorthFuzzing Then
'            Select Case arg.DataType 'the ones we dont bother to fuzz
'                Case atByte:    ary(i) = 1
'                Case atDate:    ary(i) = """Now"""
'                Case atBool:    ary(i) = "True"
'                Case atempty:   ary(i) = """"""
'                Case atSingle:  ary(i) = 1
'                Case atUnknown: ary(i) = """"""
'                Case atObject:  ary(i) = """"""
'            End Select
'        Else
'            Select Case arg.DataType 'defaults for those we do fuzz
'                Case atLong, atInteger, atdouble:    ary(i) = 1
'                Case atVariant, atstring:            ary(i) = """defaultV"""
'            End Select
'        End If
'    Next
'
'    'now that defaults are taken care of and base ary filled out..now permutate fuzzables
'    For i = 1 To Args.Count
'        Set arg = Args(i)
'
'        If Not arg.isOptional And arg.isWorthFuzzing Then
'
'            default = ary(i) 'save default val for next loop
'
'            If arg.DataType = atLong Or arg.DataType = atInteger Or arg.DataType = atdouble Then
'
'                '3 variations for numeric type data
'                ' - max positive value
'                ' - min negative value
'                ' - max negative value
'
'                Select Case arg.DataType
'                    Case atLong:    ary(i) = 2147483647
'                    Case atInteger: ary(i) = 32767
'                    Case atdouble:  ary(i) = 1.79769313486231E+308
'                End Select
'
'                WSFFile pFolder, fz, ary, fPaths, "MaxPosArg" & i
'
'                ary(i) = -1
'                WSFFile pFolder, fz, ary, fPaths, "Neg1Arg" & i
'
'                Select Case arg.DataType
'                    Case atLong:    ary(i) = -2147483647
'                    Case atInteger: ary(i) = -32768
'                    Case atdouble:  ary(i) = 3.39519326559384E-313
'                End Select
'
'                WSFFile pFolder, fz, ary, fPaths, "MaxNegArg" & i
'
'            ElseIf arg.DataType = atstring Or arg.DataType = atVariant Then
'
'                    '5 variations for string type data
'                    ' - medium long string
'                    ' - long long string
'                    ' - long long long string
'                    ' - format string test
'                    ' - try to trigger possible parsing bug
'
'                    ary(i) = "String(1050, ""A"")"
'                    WSFFile pFolder, fz, ary, fPaths, "Str1050Arg" & i
'
'                    ary(i) = "String(5001, ""A"")"
'                    WSFFile pFolder, fz, ary, fPaths, "Str5001Arg" & i
'
'                    ary(i) = "String(15000, ""A"")"
'                    WSFFile pFolder, fz, ary, fPaths, "Str15kArg" & i
'
'                    ary(i) = """%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n%n"""
'                    WSFFile pFolder, fz, ary, fPaths, "FormatStrArg" & i
'
'                    ary(i) = """\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\\\:#$%test\test\test\te?s\test\test\tes\\:#$%\ttest\test\te@st\tes\test\test\tes.\ttest\test\test\tes\test\test\te.s\ttest\test\test\tes\test\test\tes\t\\\\\\\"""
'                    WSFFile pFolder, fz, ary, fPaths, "JunkParseArg" & i
'
'            End If
'
'            ary(i) = default 'back to default so we can move to next
'
'        End If
'    Next
'
'    BuildFuzzConfigFiles = fPaths 'return filepaths
'
'End Function

'Private Function HtmlPage(ByVal pfolder As String, fz As fzConfig, ary(), fpath() As String)
'
'    Dim f As String
'    Dim tmp() As String
'    Dim argStr As String
'    Dim ik As InvokeKinds
'    Dim i As Long
'
'    If Right(pfolder, 1) <> "\" Then pfolder = pfolder & "\"
'    f = fso.GetFreeFileName(pfolder, ".html")
'
'    push tmp(), "<object classid='clsid:" & Mid(fz.objGuid, 2, Len(fz.objGuid) - 2) & "' id=target></object>"
'    push tmp(), "<script language=vbs>"
'
'    For i = 1 To UBound(ary)
'        argStr = argStr & "arg" & i & " ,"
'        If TypeName(ary(i)) = "String" Or TypeName(ary(i)) = "Variant" Then
'            push tmp(), "arg" & i & "=""" & Replace(ary(i), """", """""") & """"
'        Else
'            push tmp(), "arg" & i & "=" & ary(i)
'        End If
'    Next
'
'    ik = fz.InvokeKind
'
'    If ik = INVOKE_PROPERTYPUT Then
'        If UBound(ary) > 1 Then MsgBox "Property put with " & UBound(ary) & " args?"
'        push tmp(), "target." & fz.methodName & " = arg1"
'    Else
'        push tmp(), "target." & fz.methodName & " " & VBA.Left(argStr, Len(argStr) - 1)
'    End If
'
'    push tmp(), "</script>"
'
'    fso.WriteFile f, Join(tmp, vbCrLf)
'
'    push fpath(), f
'
'End Function

'Sub Invoke(liveInst As Object, mArgs() As Variant)
'
'    Dim vArgs() As Variant
'    Dim i As Integer, n As Integer
'    Dim VT_Flags As Integer
'    Dim a As CArgument
'
'    If UBound(mArgs) <> (Args.Count - 1) Then
'        MsgBox "Args Passed in does not match internal count!"
'        Exit Sub
'    End If
'
'    ReDim vArgs(UBound(mArgs))
'
'    For i = UBound(mArgs) To 0
'
'        'Select Case TypeName(mArgs(i))
'        '    Case "String":  VT_Flags = VT_BSTR
'        '    Case "Integer": VT_Flags = VT_BYREF Or VT_I2
'        'End Select
'
'        Set a = Args(i + 1)
'        VT_Flags = a.VT_Flags
'
'        vArgs(n) = mArgs(i)              'we have to reverse order args for call
'        CopyMemory vArgs(n), VT_Flags, 2 'modify VT struc to set our datatype
'        n = n + 1
'
'    Next
'
'    With mMemberInfo
'        tli.InvokeHookArraySub liveInst, .MemberId, .InvokeKind, vArgs
'
'
'    End With
'
'End Sub
 
 
