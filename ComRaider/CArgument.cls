VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArgument"
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

Enum argTypes
    atParamAry = 0
    atLong = 1
    atInteger = 2
    atByte = 3
    atObject = 4
    atstring = 5
    atDate = 6
    atVariant = 7
    atOther = 8
    atarray = 9
    atdouble = 10
    atempty = 11
    atBool = 12
    atSingle = 13
    atUnknown = 14
End Enum

Public isWorthFuzzing As Boolean 'is this one argument worth fuzzing
Public isDefault As Boolean
Public isOptional As Boolean
Public isByVal As Boolean
Public DataType As argTypes
Public Name As String
Public ObjectType As String

Sub ManSetType(t As argTypes, worthy As Boolean, Optional vt As TliVarType)
    
    isWorthFuzzing = worthy
    DataType = t
    
End Sub

Sub SetStdType(t)  'not done yet
    
    isWorthFuzzing = True
    
    Select Case t
        Case "Integer":     DataType = atInteger
        Case "Long":        DataType = atLong
        Case "String":      DataType = atstring
        Case "Variant":     DataType = atVariant
        Case "Double":      DataType = atdouble
        Case "Date"
                DataType = atDate
                isWorthFuzzing = False
        Case "Byte"
                DataType = atByte
                isWorthFuzzing = False
        Case "Boolean"
                DataType = atBool
                isWorthFuzzing = False
        Case "Single"
                DataType = atSingle
                isWorthFuzzing = False
        Case "Unknown"
                DataType = atUnknown
                isWorthFuzzing = False
        Case "Object"
                DataType = atObject
                isWorthFuzzing = False
        Case "Empty"
                DataType = atempty
                isWorthFuzzing = False
        Case Else
                DataType = atOther
                'MsgBox "Unknown datatype: " & t
                
    End Select

    
    
End Sub
