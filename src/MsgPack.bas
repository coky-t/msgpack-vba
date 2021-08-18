Attribute VB_Name = "MsgPack"
Option Explicit

'
' Copyright (c) 2021 Koki Takeyama
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

''
'' MessagePack for VBA
''

'
' Conditional
'

' Integer
#If Win64 Then
#Const USE_LONGLONG = True
#End If

' Integer Optimize
#Const USE_SIGNED_INT = True

' Array
#Const USE_COLLECTION = True

'
' Const
'

Private Const mpExtTimestamp As Byte = &HFF
Private Const mpExtCurrency As Byte = vbCurrency
Private Const mpExtDate As Byte = vbDate
Private Const mpExtDecimal As Byte = vbDecimal

'
' Declare
'

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As LongPtr)
#Else
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As Long)
#End If

'
' Types
'

Private Type IntegerT
    Value As Integer
End Type

Private Type LongT
    Value As Long
End Type

#If Win64 And USE_LONGLONG Then
Private Type LongLongT
    Value As LongLong
End Type
#End If

Private Type SingleT
    Value As Single
End Type

Private Type DoubleT
    Value As Double
End Type

Private Type CurrencyT
    Value As Currency
End Type

Private Type DateT
    Value As Date
End Type

Private Type Bytes2T
    Bytes(0 To 1) As Byte
End Type

Private Type Bytes4T
    Bytes(0 To 3) As Byte
End Type

Private Type Bytes8T
    Bytes(0 To 7) As Byte
End Type

''
'' MessagePack for VBA - Serialization
''

Public Function GetMPBytes( _
    Value, Optional Optimize As Boolean = True) As Byte()
    
    Select Case VarType(Value)
    
    Case vbEmpty
        GetMPBytes = GetMPBytesFromEmpty(Value, Optimize)
        
    Case vbNull
        GetMPBytes = GetMPBytesFromNull(Value, Optimize)
        
    Case vbInteger
        GetMPBytes = GetMPBytesFromInteger(Value, Optimize)
        
    Case vbLong
        GetMPBytes = GetMPBytesFromLong(Value, Optimize)
        
    Case vbSingle
        GetMPBytes = GetMPBytesFromSingle(Value, Optimize)
        
    Case vbDouble
        GetMPBytes = GetMPBytesFromDouble(Value, Optimize)
        
    Case vbCurrency
        GetMPBytes = GetMPBytesFromCurrency(Value, Optimize)
        
    Case vbDate
        GetMPBytes = GetMPBytesFromDate(Value, Optimize)
        
    Case vbString
        GetMPBytes = GetMPBytesFromString(Value, Optimize)
        
    Case vbObject
        GetMPBytes = GetMPBytesFromObject(Value, Optimize)
        
    Case vbError
        GetMPBytes = GetMPBytesFromError(Value, Optimize)
        
    Case vbBoolean
        GetMPBytes = GetMPBytesFromBoolean(Value, Optimize)
        
    Case vbVariant
        GetMPBytes = GetMPBytesFromVariant(Value, Optimize)
        
    Case vbDataObject
        GetMPBytes = GetMPBytesFromDataObject(Value, Optimize)
        
    Case vbDecimal
        GetMPBytes = GetMPBytesFromDecimal(Value, Optimize)
        
    Case vbByte
        GetMPBytes = GetMPBytesFromByte(Value, Optimize)
        
    #If Win64 And USE_LONGLONG Then
    Case vbLongLong
        GetMPBytes = GetMPBytesFromLongLong(Value, Optimize)
    #End If
        
    Case vbUserDefinedType
        GetMPBytes = GetMPBytesFromUserDefinedType(Value, Optimize)
        
    Case vbByte + vbArray
        GetMPBytes = GetMPBytesFromByteArray(Value, Optimize)
        
    Case Else
        GetMPBytes = GetMPBytesFromUnknown(Value, Optimize)
        
    End Select
End Function

'
' 0. Empty
'

Private Function GetMPBytesFromEmpty( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromEmpty = GetMPBytesFromNil
End Function

'
' 1. Null
'

Private Function GetMPBytesFromNull( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromNull = GetMPBytesFromNil
End Function

'
' 2. Integer
'

Private Function GetMPBytesFromInteger( _
    ByVal Value As Integer, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromInteger = GetMPBytesFromInt(Value)
    Else
        GetMPBytesFromInteger = GetMPBytesFromInt16(Value)
    End If
End Function

'
' 3. Long
'

Private Function GetMPBytesFromLong( _
    ByVal Value As Long, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromLong = GetMPBytesFromInt(Value)
    Else
        GetMPBytesFromLong = GetMPBytesFromInt32(Value)
    End If
End Function

'
' 4. Single
'

Private Function GetMPBytesFromSingle( _
    ByVal Value As Single, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromSingle = GetMPBytesFromFloat32(Value)
End Function

'
' 5. Double
'

Private Function GetMPBytesFromDouble( _
    ByVal Value As Double, Optional Optimize As Boolean) As Byte()
    
    ' to do - optimize
    
    GetMPBytesFromDouble = GetMPBytesFromFloat64(Value)
End Function

'
' 6. Currency
'

Private Function GetMPBytesFromCurrency( _
    ByVal Value As Currency, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromCurrency = GetMPBytesFromCur(Value)
    Else
        GetMPBytesFromCurrency = GetMPBytesFromFixExtCur8(Value)
    End If
End Function

'
' 7. Date
'

Private Function GetMPBytesFromDate( _
    ByVal Value As Date, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromDate = GetMPBytesFromTime(Value)
    Else
        GetMPBytesFromDate = GetMPBytesFromFixExtDate8(Value)
    End If
End Function

'
' 8. String
'

Private Function GetMPBytesFromString( _
    ByVal Value As String, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromString = GetMPBytesFromStr(Value)
End Function

'
' 9. Object
'

Private Function GetMPBytesFromObject( _
    Value, Optional Optimize As Boolean) As Byte()
    
    If Value Is Nothing Then
        GetMPBytesFromObject = GetMPBytesFromNil
        Exit Function
    End If
    
    Select Case TypeName(Value)
    Case "Collection"
        GetMPBytesFromObject = GetMPBytesFromCollection(Value, Optimize)
        
    Case "Dictionary"
        GetMPBytesFromObject = GetMPBytesFromDictionary(Value, Optimize)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 9. Object - Collection
'

Private Function GetMPBytesFromCollection( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromCollection = GetMPBytesFromCol(Value, Optimize)
End Function

'
' 9. Object - Dictionary
'

Private Function GetMPBytesFromDictionary( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromDictionary = GetMPBytesFromDic(Value, Optimize)
End Function

'
' 10. Error
'

Private Function GetMPBytesFromError( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Err.Raise 13 ' unmatched type
End Function

'
' 11. Boolean
'

Private Function GetMPBytesFromBoolean( _
    Value, Optional Optimize As Boolean) As Byte()
    
    If Value Then
        GetMPBytesFromBoolean = GetMPBytesFromTrue
    Else
        GetMPBytesFromBoolean = GetMPBytesFromFalse
    End If
End Function

'
' 12. Variant
'

Private Function GetMPBytesFromVariant( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Err.Raise 13 ' unmatched type
End Function

'
' 13. DataObject
'

Private Function GetMPBytesFromDataObject( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Err.Raise 13 ' unmatched type
End Function

'
' 14. Decimal
'

Private Function GetMPBytesFromDecimal( _
    Value, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromDecimal = GetMPBytesFromDec(Value)
    Else
        GetMPBytesFromDecimal = _
            GetMPBytesFromExt8(mpExtDecimal, _
                GetBytesFromDecimal(Value, True), 14)
    End If
End Function

'
' 17. Byte
'

Private Function GetMPBytesFromByte( _
    ByVal Value As Byte, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromByte = GetMPBytesFromInt(Value)
    Else
        GetMPBytesFromByte = GetMPBytesFromUInt8(Value)
    End If
End Function

'
' 20. LongLong
'

#If Win64 And USE_LONGLONG Then
Private Function GetMPBytesFromLongLong( _
    ByVal Value As LongLong, Optional Optimize As Boolean) As Byte()
    
    If Optimize Then
        GetMPBytesFromLongLong = GetMPBytesFromInt(Value)
    Else
        GetMPBytesFromLongLong = GetMPBytesFromInt64(Value)
    End If
End Function
#End If

'
' 36. UserDefinedType
'

Private Function GetMPBytesFromUserDefinedType( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Err.Raise 13 ' unmatched type
End Function

'
' 8209. (17 + 8192) Byte Array
'

Private Function GetMPBytesFromByteArray( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromByteArray = GetMPBytesFromBytes(Value)
End Function

'
' X. Unknown
'

Private Function GetMPBytesFromUnknown( _
    Value, Optional Optimize As Boolean) As Byte()
    
    If IsArray(Value) Then
        GetMPBytesFromUnknown = GetMPBytesFromArray(Value, Optimize)
    Else
        Err.Raise 13 ' unmatched type
    End If
End Function

'
' X. Unknown - Array
'

Private Function GetMPBytesFromArray( _
    Value, Optional Optimize As Boolean) As Byte()
    
    GetMPBytesFromArray = GetMPBytesFromArr(Value, Optimize)
End Function

''
'' MessagePack for VBA - Serialization - Optimize
''

'
' 2. Integer
' 3. Long
' 17. Byte
' 20. LongLong
'

Private Function GetMPBytesFromInt(Value) As Byte()
    Select Case Value
    
    #If Win64 And USE_LONGLONG Then
    Case -2147483648^ To -32769
        GetMPBytesFromInt = GetMPBytesFromInt32(Value)
    #End If
        
    Case -32768 To -129
        GetMPBytesFromInt = GetMPBytesFromInt16(Value)
        
    Case -128 To -33
        GetMPBytesFromInt = GetMPBytesFromInt8(Value)
        
    Case -32 To -1
        GetMPBytesFromInt = GetMPBytesFromNegativeFixInt(Value)
        
    Case 0 To 127 '&H7F
        GetMPBytesFromInt = GetMPBytesFromPositiveFixInt(Value)
        
    Case 128 To 255 '&H80 To &HFF
        GetMPBytesFromInt = GetMPBytesFromUInt8(Value)
        
    #If USE_SIGNED_INT Then
    Case 256 To 32767 '&H100 To &H7FFF&
        GetMPBytesFromInt = GetMPBytesFromInt16(Value)
        
    Case 32768 To 65535 '&H8000 To &HFFFF&
        GetMPBytesFromInt = GetMPBytesFromUInt16(Value)
    #Else
    Case 256 To 65535 '&H100 To &HFFFF&
        GetMPBytesFromInt = GetMPBytesFromUInt16(Value)
    #End If
        
    #If Win64 And USE_LONGLONG Then
    #If USE_SIGNED_INT Then
    Case 65536 To 2147483647 '&H10000 To &H7FFFFFFF
        GetMPBytesFromInt = GetMPBytesFromInt32(Value)
        
    Case 2147483648^ To 4294967295^ '&H80000000^ To &HFFFFFFFF^
        GetMPBytesFromInt = GetMPBytesFromUInt32(Value)
    #Else
    Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
        GetMPBytesFromInt = GetMPBytesFromUInt32(Value)
    #End If
    #End If
        
    Case Else
    #If Win64 And USE_LONGLONG Then
        GetMPBytesFromInt = GetMPBytesFromInt64(Value)
    #Else
        GetMPBytesFromInt = GetMPBytesFromInt32(Value)
    #End If
        
    End Select
End Function

'
' 6. Currency
'

Private Function GetMPBytesFromCur(ByVal Value As Currency) As Byte()
    Dim RequiredBytesCount As Long
    
    If Value >= 0@ Then
        If Value <= CCur("0.0255") Then
            RequiredBytesCount = 1
        ElseIf Value <= CCur("6.5535") Then
            RequiredBytesCount = 2
        ElseIf Value <= CCur("429496.7295") Then
            RequiredBytesCount = 4
        Else
            RequiredBytesCount = 8
        End If
    Else
        ' to do - optimize for negative value
            RequiredBytesCount = 8
    End If
    
    Select Case RequiredBytesCount
    
    Case 1
        GetMPBytesFromCur = GetMPBytesFromFixExtCur1(Value)
        
    Case 2
        GetMPBytesFromCur = GetMPBytesFromFixExtCur2(Value)
        
    Case 4
        GetMPBytesFromCur = GetMPBytesFromFixExtCur4(Value)
        
    Case Else
        GetMPBytesFromCur = GetMPBytesFromFixExtCur8(Value)
        
    End Select
End Function

'
' 7. Date
'

Private Function GetMPBytesFromTime(Value) As Byte()
    Dim RequiredBytesCount As Long
    
    If Value >= DateSerial(1970, 1, 1) Then
        If Value < DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16) Then
            RequiredBytesCount = 4
        ElseIf Value < DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4) Then
            RequiredBytesCount = 8
        Else
            RequiredBytesCount = 12
        End If
    Else
            RequiredBytesCount = 12
    End If
    
    Select Case RequiredBytesCount
    
    Case 4
        GetMPBytesFromTime = GetMPBytesFromFixExtTime4(Value)
        
    Case 8
        GetMPBytesFromTime = GetMPBytesFromFixExtTime8(Value)
        
    Case Else
        GetMPBytesFromTime = GetMPBytesFromExtTime8(Value)
        
    End Select
End Function

'
' 8. String
'

Private Function GetMPBytesFromStr(Value) As Byte()
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    If CStr(Value) = "" Then
        GetMPBytesFromStr = GetMPBytes0(&HA0)
        Exit Function
    End If
    
    Dim StrBytes() As Byte
    StrBytes = GetBytesFromString(Value)
    
    Dim StrLength As Long
    StrLength = UBound(StrBytes) - LBound(StrBytes) + 1
    
    Select Case StrLength
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &H1 To &H1F
        GetMPBytesFromStr = GetMPBytesFromFixStr(StrBytes, StrLength)
        
    'str 8           | 11011001               | 0xd9
    Case &H20 To &HFF
        GetMPBytesFromStr = GetMPBytesFromStr8(StrBytes, StrLength)
        
    'str 16          | 11011010               | 0xda
    Case &H100 To &HFFFF&
        GetMPBytesFromStr = GetMPBytesFromStr16(StrBytes, StrLength)
        
    'str 32          | 11011011               | 0xdb
    Case Else
        GetMPBytesFromStr = GetMPBytesFromStr32(StrBytes, StrLength)
        
    End Select
End Function

'
' 9. Object - Collection
'

Private Function GetMPBytesFromCol( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Select Case Value.Count
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case 0
        GetMPBytesFromCol = GetMPBytes0(&H90)
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H1 To &HF
        GetMPBytesFromCol = GetMPBytesFromFixArray(Value, Optimize)
        
    'array 16        | 11011100               | 0xdc
    Case &H10 To &HFFFF&
        GetMPBytesFromCol = GetMPBytesFromArray16(Value, Optimize)
        
    'array 32        | 11011101               | 0xdd
    Case Else
        GetMPBytesFromCol = GetMPBytesFromArray32(Value, Optimize)
        
    End Select
End Function

'
' 9. Object - Dictionary
'

Private Function GetMPBytesFromDic( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Select Case Value.Count
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case 0
        GetMPBytesFromDic = GetMPBytes0(&H80)
        
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H1 To &HF
        GetMPBytesFromDic = GetMPBytesFromFixMap(Value, Optimize)
        
    'map 16          | 11011110               | 0xde
    Case &H10 To &HFFFF&
        GetMPBytesFromDic = GetMPBytesFromMap16(Value, Optimize)
        
    'map 32          | 11011111               | 0xdf
    Case Else
        GetMPBytesFromDic = GetMPBytesFromMap32(Value, Optimize)
        
    End Select
End Function

'
' 14. Decimal
'

Private Function GetMPBytesFromDec(Value) As Byte()
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromDecimal(Value, True)
    
    Dim RequiredBytesCount As Long
    
    If (BytesBE(0) = 0) And (BytesBE(1) = 0) Then
        ' positive value, no scaling
        If Value <= CDec("255") Then
            RequiredBytesCount = 1
        ElseIf Value <= CDec("65535") Then
            RequiredBytesCount = 2
        ElseIf Value <= CDec("4294967295") Then
            RequiredBytesCount = 4
        ElseIf Value <= CDec("18446744073709551615") Then
            RequiredBytesCount = 8
        ElseIf Value <= CDec("79228162514264337593543950335") Then
            RequiredBytesCount = 12
        Else
            RequiredBytesCount = 14
        End If
    Else
        ' to do - optimize for negative value or scaling
            RequiredBytesCount = 14
    End If
    
    Select Case RequiredBytesCount
    
    Case 1
        GetMPBytesFromDec = GetMPBytesFromFixExtDec1(Value)
        
    Case 2
        GetMPBytesFromDec = GetMPBytesFromFixExtDec2(Value)
        
    Case 4
        GetMPBytesFromDec = GetMPBytesFromFixExtDec4(Value)
        
    Case 8
        GetMPBytesFromDec = GetMPBytesFromFixExtDec8(Value)
        
    Case 12
        GetMPBytesFromDec = GetMPBytesFromExtDec8_12(Value)
        
    Case Else
        GetMPBytesFromDec = GetMPBytesFromExtDec8_14(Value)
        
    End Select
End Function

'
' 8209. (17 + 8192) Byte Array
'

Private Function GetMPBytesFromBytes(Value) As Byte()
    Dim Length As Long
    
    On Error Resume Next
    Length = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
    
    Select Case Length
    
    'bin 8           | 11000100               | 0xc4
    Case &H0
        GetMPBytesFromBytes = GetMPBytes1A(&HC4, 0)
        
    'bin 8           | 11000100               | 0xc4
    Case &H1 To &HFF
        GetMPBytesFromBytes = GetMPBytesFromBin8(Value, Length)
        
    'bin 16          | 11000101               | 0xc5
    Case &H100 To &HFFFF&
        GetMPBytesFromBytes = GetMPBytesFromBin16(Value, Length)
        
    'bin 32          | 11000110               | 0xc6
    Case Else
        GetMPBytesFromBytes = GetMPBytesFromBin32(Value, Length)
        
    End Select
End Function

'
' X. Unknown - Array
'

Private Function GetMPBytesFromArr( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Dim Length As Long
    
    On Error Resume Next
    Length = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
    
    Select Case Length
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case 0
        GetMPBytesFromArr = GetMPBytes0(&H90)
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H1 To &HF
        GetMPBytesFromArr = GetMPBytesFromFixArray(Value, Optimize)
        
    'array 16        | 11011100               | 0xdc
    Case &H10 To &HFFFF&
        GetMPBytesFromArr = GetMPBytesFromArray16(Value, Optimize)
        
    'array 32        | 11011101               | 0xdd
    Case Else
        GetMPBytesFromArr = GetMPBytesFromArray32(Value, Optimize)
        
    End Select
End Function

''
'' MessagePack for VBA - Serialization - Extension
''

'
' -1. Timestamp
'

'fixext 4        | 11010110               | 0xd6
'timestamp 32 stores the number of seconds that have elapsed since 1970-01-01 00:00:00 UTC
'in an 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |   -1   |   seconds in 32-bit unsigned int  |
'+--------+--------+--------+--------+--------+--------+
'* Timestamp 32 format can represent a timestamp in [1970-01-01 00:00:00 UTC, 2106-02-07 06:28:16 UTC) range. Nanoseconds part is 0.

Private Function GetMPBytesFromFixExtTime4(ByVal DateTime As Date) As Byte()
    Debug.Assert (DateTime >= DateSerial(1970, 1, 1))
    Debug.Assert (DateTime < DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16))
    
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    GetMPBytesFromFixExtTime4 = _
        GetMPBytesFromFixExt4( _
            mpExtTimestamp, GetBytesFromUInt32(Seconds, True))
End Function

'fixext 8        | 11010111               | 0xd7
'timestamp 64 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 32-bit unsigned integers:
'+--------+--------+--------+--------+--------+------|-+--------+--------+--------+--------+
'|  0xd7  |   -1   | nanosec. in 30-bit unsigned int |   seconds in 34-bit unsigned int    |
'+--------+--------+--------+--------+--------+------^-+--------+--------+--------+--------+
'* Timestamp 64 format can represent a timestamp in [1970-01-01 00:00:00.000000000 UTC, 2514-05-30 01:53:04.000000000 UTC) range.

Private Function GetMPBytesFromFixExtTime8(ByVal DateTime As Date) As Byte()
    Debug.Assert (DateTime >= DateSerial(1970, 1, 1))
    Debug.Assert (DateTime < DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4))
    
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    GetMPBytesFromFixExtTime8 = _
        GetMPBytesFromFixExt8( _
            mpExtTimestamp, GetBytesFromUInt64(CDec(Seconds), True))
End Function

'ext 8           | 11000111               | 0xc7
'timestamp 96 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 64-bit signed integer and 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+--------+
'|  0xc7  |   12   |   -1   |nanoseconds in 32-bit unsigned int |
'+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                    seconds in 64-bit signed int                        |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* Timestamp 96 format can represent a timestamp in [-292277022657-01-27 08:29:52 UTC, 292277026596-12-04 15:30:08.000000000 UTC) range.
'* In timestamp 64 and timestamp 96 formats, nanoseconds must not be larger than 999999999.

Private Function GetMPBytesFromExtTime8(ByVal DateTime As Date) As Byte()
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    Dim BytesBE8() As Byte
    BytesBE8 = GetBytesFromInt64(Seconds, True)
    
    Dim BytesBE12(0 To 11) As Byte
    CopyBytes BytesBE12, 4, BytesBE8, 0, 8
    
    GetMPBytesFromExtTime8 = GetMPBytesFromExt8(mpExtTimestamp, BytesBE12, 12)
End Function

'
' 6. Currency
'

Private Function GetMPBytesFromFixExtCur1(ByVal Value As Currency) As Byte()
    Debug.Assert ((Value >= 0@) And (Value <= CCur("0.0255")))
    
    Dim BytesBE8() As Byte
    BytesBE8 = GetBytesFromCurrency(Value, True)
    
    Dim BytesBE1(0)
    BytesBE1(0) = BytesBE8(7)
    
    GetMPBytesFromFixExtCur1 = GetMPBytesFromFixExt1(mpExtCurrency, BytesBE1)
End Function

Private Function GetMPBytesFromFixExtCur2(ByVal Value As Currency) As Byte()
    Debug.Assert ((Value >= 0@) And (Value <= CCur("6.5535")))
    
    Dim BytesBE8() As Byte
    BytesBE8 = GetBytesFromCurrency(Value, True)
    
    Dim BytesBE2(0 To 1) As Byte
    CopyBytes BytesBE2, 0, BytesBE8, 6, 2
    
    GetMPBytesFromFixExtCur2 = GetMPBytesFromFixExt2(mpExtCurrency, BytesBE2)
End Function

Private Function GetMPBytesFromFixExtCur4(ByVal Value As Currency) As Byte()
    Debug.Assert ((Value >= 0@) And (Value <= CCur("429496.7295")))
    
    Dim BytesBE8() As Byte
    BytesBE8 = GetBytesFromCurrency(Value, True)
    
    Dim BytesBE4(0 To 3) As Byte
    CopyBytes BytesBE4, 0, BytesBE8, 4, 4
    
    GetMPBytesFromFixExtCur4 = GetMPBytesFromFixExt4(mpExtCurrency, BytesBE4)
End Function

Private Function GetMPBytesFromFixExtCur8(ByVal Value As Currency) As Byte()
    Dim BytesBE8() As Byte
    BytesBE8 = GetBytesFromCurrency(Value, True)
    
    GetMPBytesFromFixExtCur8 = GetMPBytesFromFixExt8(mpExtCurrency, BytesBE8)
End Function

'
' 7. Date
'

Private Function GetMPBytesFromFixExtDate8(ByVal Value As Date) As Byte()
    GetMPBytesFromFixExtDate8 = _
        GetMPBytesFromFixExt8(mpExtDate, GetBytesFromDate(Value, True))
End Function

'
' 14. Decimal
'

Private Function GetMPBytesFromFixExtDec1(ByVal Value As Variant) As Byte()
    Debug.Assert ((Value >= CDec(0)) And (Value <= CDec("255")))
    
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    Dim BytesBE1(0) As Byte
    BytesBE1(0) = BytesBE14(13)
    
    GetMPBytesFromFixExtDec1 = GetMPBytesFromFixExt1(mpExtDecimal, BytesBE1)
End Function

Private Function GetMPBytesFromFixExtDec2(ByVal Value As Variant) As Byte()
    Debug.Assert ((Value >= CDec(0)) And (Value <= CDec("65535")))
    
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    Dim BytesBE2(0 To 1) As Byte
    CopyBytes BytesBE2, 0, BytesBE14, 12, 2
    
    GetMPBytesFromFixExtDec2 = GetMPBytesFromFixExt2(mpExtDecimal, BytesBE2)
End Function

Private Function GetMPBytesFromFixExtDec4(ByVal Value As Variant) As Byte()
    Debug.Assert ((Value >= CDec(0)) And (Value <= CDec("4294967295")))
    
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    Dim BytesBE4(0 To 3) As Byte
    CopyBytes BytesBE4, 0, BytesBE14, 10, 4
    
    GetMPBytesFromFixExtDec4 = GetMPBytesFromFixExt4(mpExtDecimal, BytesBE4)
End Function

Private Function GetMPBytesFromFixExtDec8(ByVal Value As Variant) As Byte()
    Debug.Assert _
        ((Value >= CDec(0)) And (Value <= CDec("18446744073709551615")))
    
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    Dim BytesBE8(0 To 7) As Byte
    CopyBytes BytesBE8, 0, BytesBE14, 6, 8
    
    GetMPBytesFromFixExtDec8 = GetMPBytesFromFixExt8(mpExtDecimal, BytesBE8)
End Function

Private Function GetMPBytesFromExtDec8_12(ByVal Value As Variant) As Byte()
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    Dim BytesBE12(0 To 11) As Byte
    CopyBytes BytesBE12, 0, BytesBE14, 2, 12
    
    GetMPBytesFromExtDec8_12 = GetMPBytesFromExt8(mpExtDecimal, BytesBE12, 12)
End Function

Private Function GetMPBytesFromExtDec8_14(ByVal Value As Variant) As Byte()
    Dim BytesBE14() As Byte
    BytesBE14 = GetBytesFromDecimal(Value, True)
    
    GetMPBytesFromExtDec8_14 = GetMPBytesFromExt8(mpExtDecimal, BytesBE14, 14)
End Function

''
'' MessagePack for VBA - Serialization - Core
''

'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
'positive fixint stores 7-bit positive integer
'+--------+
'|0XXXXXXX|
'+--------+
'* 0XXXXXXX is 8-bit unsigned integer

Private Function GetMPBytesFromPositiveFixInt(ByVal Value As Byte) As Byte()
    Debug.Assert (Value <= &H7F)
    GetMPBytesFromPositiveFixInt = GetMPBytes0(Value)
End Function

'fixmap          | 1000xxxx               | 0x80 - 0x8f
'fixmap stores a map whose length is upto 15 elements
'+--------+~~~~~~~~~~~~~~~~~+
'|1000XXXX|   N*2 objects   |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetMPBytesFromFixMap( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert (Value.Count <= &HF)
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &H80 Or Value.Count
    
    AddMPBytesFromMap MPBytes, Value, Optimize
    
    GetMPBytesFromFixMap = MPBytes
End Function

'fixarray        | 1001xxxx               | 0x90 - 0x9f
'fixarray stores an array whose length is upto 15 elements:
'+--------+~~~~~~~~~~~~~~~~~+
'|1001XXXX|    N objects    |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of an array

Private Function GetMPBytesFromFixArray( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert (Count <= &HF)
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &H90 Or Count
    
    AddMPBytesFromArray MPBytes, Value, Optimize
    
    GetMPBytesFromFixArray = MPBytes
End Function

'fixstr          | 101xxxxx               | 0xa0 - 0xbf
'fixstr stores a byte array whose length is upto 31 bytes:
'+--------+========+
'|101XXXXX|  data  |
'+--------+========+
'* XXXXX is a 5-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetMPBytesFromFixStr( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert ((StrLength > 0) And (StrLength <= &H1F))
    
    GetMPBytesFromFixStr = GetMPBytes1B(&HA0 Or StrLength, StrBytes)
End Function

'nil             | 11000000               | 0xc0
'nil:
'+--------+
'|  0xc0  |
'+--------+

Private Function GetMPBytesFromNil() As Byte()
    GetMPBytesFromNil = GetMPBytes0(&HC0)
End Function

'false           | 11000010               | 0xc2
'false:
'+--------+
'|  0xc2  |
'+--------+

Private Function GetMPBytesFromFalse() As Byte()
    GetMPBytesFromFalse = GetMPBytes0(&HC2)
End Function

'true            | 11000011               | 0xc3
'true:
'+--------+
'|  0xc3  |
'+--------+

Private Function GetMPBytesFromTrue() As Byte()
    GetMPBytesFromTrue = GetMPBytes0(&HC3)
End Function

'bin 8           | 11000100               | 0xc4
'bin 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xc4  |XXXXXXXX|  data  |
'+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is the length of data

Private Function GetMPBytesFromBin8( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromBin8 = GetMPBytes2A(&HC4, BinLength, BinBytes)
End Function

'bin 16          | 11000101               | 0xc5
'bin 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xc5  |YYYYYYYY|YYYYYYYY|  data  |
'+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data

Private Function GetMPBytesFromBin16( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromBin16 = _
        GetMPBytes2B(&HC5, GetBytesFromUInt16(BinLength, True), BinBytes)
End Function

'bin 32          | 11000110               | 0xc6
'bin 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xc6  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data

Private Function GetMPBytesFromBin32( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromBin32 = _
        GetMPBytes2B(&HC6, GetBytesFromUInt32(BinLength, True), BinBytes)
End Function

'ext 8           | 11000111               | 0xc7
'ext 8 stores an integer and a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+--------+========+
'|  0xc7  |XXXXXXXX|  type  |  data  |
'+--------+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromExt8( _
    ExtType As Byte, BinBytes, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromExt8 = GetMPBytes3A(&HC7, BinLength, ExtType, BinBytes)
End Function

'ext 16          | 11001000               | 0xc8
'ext 16 stores an integer and a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+--------+========+
'|  0xc8  |YYYYYYYY|YYYYYYYY|  type  |  data  |
'+--------+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromExt16( _
    ExtType As Byte, BinBytes, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromExt16 = GetMPBytes3B(&HC8, _
        GetBytesFromUInt16(BinLength, True), ExtType, BinBytes)
End Function

'ext 32          | 11001001               | 0xc9
'ext 32 stores an integer and a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+--------+========+
'|  0xc9  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  type  |  data  |
'+--------+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a big-endian 32-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromExt32( _
    ExtType As Byte, BinBytes, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetMPBytesFromExt32 = GetMPBytes3B(&HC9, _
        GetBytesFromUInt32(BinLength, True), ExtType, BinBytes)
End Function

'float 32        | 11001010               | 0xca
'float 32 stores a floating point number in IEEE 754 single precision floating point number format:
'+--------+--------+--------+--------+--------+
'|  0xca  |XXXXXXXX|XXXXXXXX|XXXXXXXX|XXXXXXXX|
'+--------+--------+--------+--------+--------+
'* XXXXXXXX_XXXXXXXX_XXXXXXXX_XXXXXXXX is a big-endian IEEE 754 single precision floating point number.
'  Extension of precision from single-precision to double-precision does not lose precision.

Private Function GetMPBytesFromFloat32(ByVal Value As Single) As Byte()
    GetMPBytesFromFloat32 = _
        GetMPBytes1B(&HCA, GetBytesFromFloat32(Value, True))
End Function

'float 64        | 11001011               | 0xcb
'float 64 stores a floating point number in IEEE 754 double precision floating point number format:
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcb  |YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY is a big-endian
'  IEEE 754 double precision floating point number

Private Function GetMPBytesFromFloat64(ByVal Value As Double) As Byte()
    GetMPBytesFromFloat64 = _
        GetMPBytes1B(&HCB, GetBytesFromFloat64(Value, True))
End Function

'uint 8          | 11001100               | 0xcc
'uint 8 stores a 8-bit unsigned integer
'+--------+--------+
'|  0xcc  |ZZZZZZZZ|
'+--------+--------+

Private Function GetMPBytesFromUInt8(ByVal Value As Byte) As Byte()
    GetMPBytesFromUInt8 = GetMPBytes1A(&HCC, Value)
End Function

'uint 16         | 11001101               | 0xcd
'uint 16 stores a 16-bit big-endian unsigned integer
'+--------+--------+--------+
'|  0xcd  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+

Private Function GetMPBytesFromUInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    GetMPBytesFromUInt16 = GetMPBytes1B(&HCD, GetBytesFromUInt16(Value, True))
End Function

'uint 32         | 11001110               | 0xce
'uint 32 stores a 32-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+
'|  0xce  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+

Private Function GetMPBytesFromUInt32(ByVal Value) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    GetMPBytesFromUInt32 = GetMPBytes1B(&HCE, GetBytesFromUInt32(Value, True))
End Function

'uint 64         | 11001111               | 0xcf
'uint 64 stores a 64-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+

Private Function GetMPBytesFromUInt64(ByVal Value) As Byte()
    Debug.Assert (Value >= 0)
    GetMPBytesFromUInt64 = GetMPBytes1B(&HCF, GetBytesFromUInt64(Value, True))
End Function

'int 8           | 11010000               | 0xd0
'int 8 stores a 8-bit signed integer
'+--------+--------+
'|  0xd0  |ZZZZZZZZ|
'+--------+--------+

Private Function GetMPBytesFromInt8(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -128) And (Value <= &H7F))
    GetMPBytesFromInt8 = GetMPBytes1A(&HD0, Value And &HFF)
End Function

'int 16          | 11010001               | 0xd1
'int 16 stores a 16-bit big-endian signed integer
'+--------+--------+--------+
'|  0xd1  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+

Private Function GetMPBytesFromInt16(ByVal Value As Integer) As Byte()
    GetMPBytesFromInt16 = GetMPBytes1B(&HD1, GetBytesFromInt16(Value, True))
End Function

'int 32          | 11010010               | 0xd2
'int 32 stores a 32-bit big-endian signed integer
'+--------+--------+--------+--------+--------+
'|  0xd2  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+

Private Function GetMPBytesFromInt32(ByVal Value As Long) As Byte()
    GetMPBytesFromInt32 = GetMPBytes1B(&HD2, GetBytesFromInt32(Value, True))
End Function

'int 64          | 11010011               | 0xd3
'int 64 stores a 64-bit big-endian signed integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd3  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+

Private Function GetMPBytesFromInt64(ByVal Value) As Byte()
    GetMPBytesFromInt64 = GetMPBytes1B(&HD3, GetBytesFromInt64(Value, True))
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromFixExt1(ExtType As Byte, BinBytes) As Byte()
    GetMPBytesFromFixExt1 = GetMPBytes2A(&HD4, ExtType, BinBytes)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromFixExt2(ExtType As Byte, BinBytes) As Byte()
    GetMPBytesFromFixExt2 = GetMPBytes2A(&HD5, ExtType, BinBytes)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromFixExt4(ExtType As Byte, BinBytes) As Byte()
    GetMPBytesFromFixExt4 = GetMPBytes2A(&HD6, ExtType, BinBytes)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromFixExt8(ExtType As Byte, BinBytes) As Byte()
    GetMPBytesFromFixExt8 = GetMPBytes2A(&HD7, ExtType, BinBytes)
End Function

'fixext 16       | 11011000               | 0xd8
'fixext 16 stores an integer and a byte array whose length is 16 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd8  |  type  |                                  data
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                              data (cont.)                              |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetMPBytesFromFixExt16(ExtType As Byte, BinBytes) As Byte()
    GetMPBytesFromFixExt16 = GetMPBytes2A(&HD8, ExtType, BinBytes)
End Function

'str 8           | 11011001               | 0xd9
'str 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xd9  |YYYYYYYY|  data  |
'+--------+--------+========+
'* YYYYYYYY is a 8-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetMPBytesFromStr8( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetMPBytesFromStr8 = GetMPBytes2A(&HD9, StrLength, StrBytes)
End Function

'str 16          | 11011010               | 0xda
'str 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xda  |ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetMPBytesFromStr16( _
    StrBytes() As Byte, ByVal StrLength As Long) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetMPBytesFromStr16 = _
        GetMPBytes2B(&HDA, GetBytesFromUInt16(StrLength, True), StrBytes)
End Function

'str 32          | 11011011               | 0xdb
'str 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xdb  |AAAAAAAA|AAAAAAAA|AAAAAAAA|AAAAAAAA|  data  |
'+--------+--------+--------+--------+--------+========+
'* AAAAAAAA_AAAAAAAA_AAAAAAAA_AAAAAAAA is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetMPBytesFromStr32( _
    StrBytes() As Byte, ByVal StrLength) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetMPBytesFromStr32 = _
        GetMPBytes2B(&HDB, GetBytesFromUInt32(StrLength, True), StrBytes)
End Function

'array 16        | 11011100               | 0xdc
'array 16 stores an array whose length is upto (2^16)-1 elements:
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdc  |YYYYYYYY|YYYYYYYY|    N objects    |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of an array

Private Function GetMPBytesFromArray16( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert (Count <= &HFFFF&)
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &HDC
    AddBytes MPBytes, GetBytesFromUInt16(Count, True)
    
    AddMPBytesFromArray MPBytes, Value, Optimize
    
    GetMPBytesFromArray16 = MPBytes
End Function

'array 32        | 11011101               | 0xdd
'array 32 stores an array whose length is upto (2^32)-1 elements:
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdd  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|    N objects    |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of an array

Private Function GetMPBytesFromArray32( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &HDD
    AddBytes MPBytes, GetBytesFromUInt32(Count, True)
    
    AddMPBytesFromArray MPBytes, Value, Optimize
    
    GetMPBytesFromArray32 = MPBytes
End Function

'map 16          | 11011110               | 0xde
'map 16 stores a map whose length is upto (2^16)-1 elements
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xde  |YYYYYYYY|YYYYYYYY|   N*2 objects   |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetMPBytesFromMap16( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Debug.Assert (Value.Count <= &HFFFF&)
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &HDE
    AddBytes MPBytes, GetBytesFromUInt16(Value.Count, True)
    
    AddMPBytesFromMap MPBytes, Value, Optimize
    
    GetMPBytesFromMap16 = MPBytes
End Function

'map 32          | 11011111               | 0xdf
'map 32 stores a map whose length is upto (2^32)-1 elements
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|   N*2 objects   |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetMPBytesFromMap32( _
    Value, Optional Optimize As Boolean) As Byte()
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0)
    MPBytes(0) = &HDF
    AddBytes MPBytes, GetBytesFromUInt32(Value.Count, True)
    
    AddMPBytesFromMap MPBytes, Value, Optimize
    
    GetMPBytesFromMap32 = MPBytes
End Function

'negative fixint | 111xxxxx               | 0xe0 - 0xff
'negative fixint stores 5-bit negative integer
'+--------+
'|111YYYYY|
'+--------+
'* 111YYYYY is 8-bit signed integer

Private Function GetMPBytesFromNegativeFixInt(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -32) And (Value < 0))
    GetMPBytesFromNegativeFixInt = GetMPBytes0(Value And &HFF)
End Function

''
'' MessagePack for VBA - Serialization - Map Helper
''

Private Sub AddMPBytesFromMap( _
    MPBytes() As Byte, Value, Optional Optimize As Boolean)
    
    Debug.Assert (TypeName(Value) = "Dictionary")
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        AddBytes MPBytes, GetMPBytes(Keys(Index), Optimize)
        AddBytes MPBytes, GetMPBytes(Value.Item(Keys(Index)), Optimize)
    Next
End Sub

''
'' MessagePack for VBA - Serialization - Array Helper
''

Private Sub AddMPBytesFromArray( _
    MPBytes() As Byte, Value, Optional Optimize As Boolean)
    
    Dim LB As Long
    Dim UB As Long
    
    If IsArray(Value) Then
        LB = LBound(Value)
        UB = UBound(Value)
        
    ElseIf TypeName(Value) = "Collection" Then
        LB = 1
        UB = Value.Count
        
    Else
        Err.Raise 13 ' unmatched type
        
    End If
    
    Dim Index As Long
    For Index = LB To UB
        AddBytes MPBytes, GetMPBytes(Value(Index), Optimize)
    Next
End Sub

''
'' MessagePack for VBA - Serialization - Formatter
''

Private Function GetMPBytes0(FormatValue As Byte) As Byte()
    Dim MPBytes(0) As Byte
    MPBytes(0) = FormatValue
    GetMPBytes0 = MPBytes
End Function

Private Function GetMPBytes1A(FormatValue As Byte, SrcByte As Byte) As Byte()
    Dim MPBytes(0 To 1) As Byte
    MPBytes(0) = FormatValue
    MPBytes(1) = SrcByte
    GetMPBytes1A = MPBytes
End Function

Private Function GetMPBytes1B( _
    FormatValue As Byte, SrcBytes() As Byte) As Byte()
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    
    Dim SrcLen As Long
    SrcLen = SrcUB - SrcLB + 1
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0 To SrcLen)
    MPBytes(0) = FormatValue
    
    CopyBytes MPBytes, 1, SrcBytes, SrcLB, SrcLen
    
    GetMPBytes1B = MPBytes
End Function

Private Function GetMPBytes2A( _
    FormatValue As Byte, SrcByte1 As Byte, SrcBytes2) As Byte()
    'FormatValue As Byte, SrcByte1 As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0 To 1 + SrcLen2)
    MPBytes(0) = FormatValue
    MPBytes(1) = SrcByte1
    
    CopyBytes MPBytes, 2, SrcBytes2, SrcLB2, SrcLen2
    
    GetMPBytes2A = MPBytes
End Function

Private Function GetMPBytes2B( _
    FormatValue As Byte, SrcBytes1, SrcBytes2) As Byte()
    'FormatValue As Byte, SrcBytes1() As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0 To SrcLen1 + SrcLen2)
    MPBytes(0) = FormatValue
    
    CopyBytes MPBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    CopyBytes MPBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    GetMPBytes2B = MPBytes
End Function

Private Function GetMPBytes3A(FormatValue As Byte, _
    SrcByte1 As Byte, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0 To 1 + 1 + SrcLen3)
    MPBytes(0) = FormatValue
    MPBytes(1) = SrcByte1
    MPBytes(2) = SrcByte2
    
    CopyBytes MPBytes, 3, SrcBytes3, SrcLB3, SrcLen3
    
    GetMPBytes3A = MPBytes
End Function

Private Function GetMPBytes3B(FormatValue As Byte, _
    SrcBytes1, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim MPBytes() As Byte
    ReDim MPBytes(0 To SrcLen1 + 1 + SrcLen3)
    Bytes(0) = FormatValue
    
    CopyBytes MPBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    
    MPBytes(SrcLen1 + 1) = SrcByte2
    
    CopyBytes MPBytes, 1 + SrcLen1 + 1, SrcBytes3, SrcLB3, SrcLen3
    
    GetMPBytes3B = MPBytes
End Function

''
'' MessagePack for VBA - Serialization - Bytes Operator
''

Private Sub AddBytes(DstBytes() As Byte, SrcBytes() As Byte)
    Dim DstLB As Long
    Dim DstUB As Long
    DstLB = LBound(DstBytes)
    DstUB = UBound(DstBytes)
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    Dim SrcLen As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    SrcLen = SrcUB - SrcLB + 1
    
    ReDim Preserve DstBytes(DstLB To DstUB + SrcLen)
    CopyBytes DstBytes, DstUB + 1, SrcBytes, SrcLB, SrcLen
End Sub

Private Sub CopyBytes( _
    DstBytes() As Byte, DstIndex As Long, _
    SrcBytes, SrcIndex As Long, ByVal Length As Long)
    'SrcBytes() As Byte, SrcIndex As Long, ByVal Length As Long)
    
    Dim Offset As Long
    For Offset = 0 To Length - 1
        DstBytes(DstIndex + Offset) = SrcBytes(SrcIndex + Offset)
    Next
End Sub

Private Sub ReverseBytes( _
    ByRef Bytes() As Byte, _
    Optional Index As Long, _
    Optional ByVal Length As Long)
    
    Dim UB As Long
    
    If Length = 0 Then
        UB = UBound(Bytes)
        Length = UB - Index + 1
    Else
        UB = Index + Length - 1
    End If
    
    Dim Offset As Long
    For Offset = 0 To (Length \ 2) - 1
        Dim Temp As Byte
        Temp = Bytes(Index + Offset)
        Bytes(Index + Offset) = Bytes(UB - Offset)
        Bytes(UB - Offset) = Temp
    Next
End Sub

''
'' MessagePack for VBA - Serialization - Converter
''

'
' CA. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromFloat32( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat32 = GetBytesFromSingle(Value, BigEndian)
End Function

'
' CB. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromFloat64( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat64 = GetBytesFromDouble(Value, BigEndian)
End Function

'
' CD. UInt16 - a 16-bit unsigned integer
'

Private Function GetBytesFromUInt16( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    
    Dim Bytes4() As Byte
    Bytes4 = GetBytesFromLong(Value, BigEndian)
    
    Dim Bytes(0 To 1) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes4, 2, 2
    Else
        CopyBytes Bytes, 0, Bytes4, 0, 2
    End If
    
    GetBytesFromUInt16 = Bytes
End Function

'
' CE. UInt32 - a 32-bit unsigned integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromUInt32( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    
    Dim Bytes8() As Byte
    Bytes8 = GetBytesFromLongLong(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes8, 4, 4
    Else
        CopyBytes Bytes, 0, Bytes8, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#Else

Private Function GetBytesFromUInt32( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("4294967295")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 10, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#End If

'
' CF. UInt64 - a 64-bit unsigned integer
'

Private Function GetBytesFromUInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("18446744073709551615")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 7) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 6, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 8
    End If
    
    GetBytesFromUInt64 = Bytes
End Function

'
' D0. Int8 - a 8-bit signed integer
'

Private Function GetBytesFromInt8( _
    ByVal Value As Integer) As Byte()
    
    Debug.Assert ((Value >= -128) And (Value <= &H7F))
    
    Dim Bytes(0) As Byte
    Bytes(0) = Value And &HFF
    
    GetBytesFromInt8 = Bytes
End Function

'
' D1. Int16 - a 16-bit signed integer
'

Private Function GetBytesFromInt16( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt16 = GetBytesFromInteger(Value, BigEndian)
End Function

'
' D2. Int32 - a 32-bit signed integer
'

Private Function GetBytesFromInt32( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt32 = GetBytesFromLong(Value, BigEndian)
End Function

'
' D3. Int64 - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromInt64( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt64 = GetBytesFromLongLong(Value, BigEndian)
End Function

#Else

Private Function GetBytesFromInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert _
        ((Value >= CDec("-9223372036854775808")) And _
        (Value <= CDec("9223372036854775807")))
    
    Dim Bytes14() As Byte
    Dim Bytes(0 To 7) As Byte
    Dim Offset As Long
    
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        If Value < 0 Then
            Bytes14 = GetBytesFromDecimal(Value + 1, BigEndian)
            For Offset = 0 To 7
                Bytes(0 + Offset) = Not Bytes14(6 + Offset)
            Next
        Else
            Bytes14 = GetBytesFromDecimal(Value, BigEndian)
            CopyBytes Bytes, 0, Bytes14, 6, 8
        End If
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        If Value < 0 Then
            Bytes14 = GetBytesFromDecimal(Value + 1, BigEndian)
            For Offset = 0 To 7
                Bytes(0 + Offset) = Not Bytes14(0 + Offset)
            Next
        Else
            Bytes14 = GetBytesFromDecimal(Value, BigEndian)
            CopyBytes Bytes, 0, Bytes14, 0, 8
        End If
    End If
    
    GetBytesFromInt64 = Bytes
End Function

#End If

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetBytesFromInteger( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    Dim I As IntegerT
    I.Value = Value
    
    Dim B2 As Bytes2T
    LSet B2 = I
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    GetBytesFromInteger = B2.Bytes
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetBytesFromLong( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Dim L As LongT
    L.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = L
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromLong = B4.Bytes
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromSingle( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    Dim S As SingleT
    S.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = S
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromSingle = B4.Bytes
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromDouble( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DoubleT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDouble = B8.Bytes
End Function

'
' 6. Currency - a 64-bit number
'in an integer format, scaled by 10,000 to give a fixed-point number
'with 15 digits to the left of the decimal point and 4 digits to the right.
'

Private Function GetBytesFromCurrency( _
    ByVal Value As Currency, Optional BigEndian As Boolean) As Byte()
    
    Dim C As CurrencyT
    C.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = C
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromCurrency = B8.Bytes
End Function

'
' 7. Date - an IEEE 754 double precision floating point number
'that represent dates ranging from 1 January 100, to 31 December 9999,
'and times from 0:00:00 to 23:59:59.
'
'When other numeric types are converted to Date,
'values to the left of the decimal represent date information,
'while values to the right of the decimal represent time.
'Midnight is 0 and midday is 0.5.
'

Private Function GetBytesFromDate( _
    ByVal Value As Date, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DateT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDate = B8.Bytes
End Function

'
' 8. String
'

Private Function GetBytesFromString( _
    ByVal Value As String, Optional Charset As String = "utf-8") As Byte()
    
    Debug.Assert (Value <> "")
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        .WriteText Value
        
        .Position = 0
        .Type = 1 'ADODB.adTypeBinary
        If Charset = "utf-8" Then
            .Position = 3 ' avoid BOM
        End If
        GetBytesFromString = .Read
        
        .Close
    End With
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetBytesFromDecimal( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    CopyMemory ByVal VarPtr(BytesRaw(0)), ByVal VarPtr(Value), 16
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    CopyBytes BytesX, 0, BytesRaw, 2, 14
    
    ' Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim Bytes(0 To 13) As Byte
    
    ' data low bytes
    CopyBytes Bytes, 0, BytesX, 6, 8
    
    ' data high bytes
    CopyBytes Bytes, 8, BytesX, 2, 4
    
    ' scale
    Bytes(12) = BytesX(0)
    
    ' sign
    Bytes(13) = BytesX(1)
    
    If BigEndian Then
        ReverseBytes Bytes
    End If
    
    GetBytesFromDecimal = Bytes
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromLongLong( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Dim LL As LongLongT
    LL.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = LL
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromLongLong = B8.Bytes
End Function

#End If

''
'' MessagePack for VBA - Deserialization
''

Private Function GetMPLength( _
    MPBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case MPBytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        GetMPLength = 1
        
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        ItemCount = MPBytes(Index) And &HF
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount * 2, MPBytes, Index + 1)
        GetMPLength = 1 + ItemLength
        
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        ItemCount = MPBytes(Index) And &HF
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount, MPBytes, Index + 1)
        GetMPLength = 1 + ItemLength
        
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        ItemLength = (MPBytes(Index) And &H1F)
        GetMPLength = 1 + ItemLength
        
    'nil             | 11000000               | 0xc0
    Case &HC0
        GetMPLength = 1
        
    '(never used)    | 11000001               | 0xc1
    Case &HC1
        GetMPLength = 1
        
    'false           | 11000010               | 0xc2
    Case &HC2
        GetMPLength = 1
        
    'true            | 11000011               | 0xc3
    Case &HC3
        GetMPLength = 1
        
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        ItemLength = MPBytes(Index + 1)
        GetMPLength = 1 + 1 + ItemLength
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        ItemLength = GetUInt16FromBytes(MPBytes, Index + 1, True)
        GetMPLength = 1 + 2 + ItemLength
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        ItemLength = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
        GetMPLength = 1 + 4 + ItemLength
        
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        ItemLength = MPBytes(Index + 1)
        GetMPLength = 1 + 1 + 1 + ItemLength
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        ItemLength = GetUInt16FromBytes(MPBytes, Index + 1, True)
        GetMPLength = 1 + 2 + 1 + ItemLength
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        ItemLength = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
        GetMPLength = 1 + 4 + 1 + ItemLength
        
    'float 32        | 11001010               | 0xca
    Case &HCA
        GetMPLength = 1 + 4
        
    'float 64        | 11001011               | 0xcb
    Case &HCB
        GetMPLength = 1 + 8
        
    'uint 8          | 11001100               | 0xcc
    Case &HCC
        GetMPLength = 1 + 1
        
    'uint 16         | 11001101               | 0xcd
    Case &HCD
        GetMPLength = 1 + 2
        
    'uint 32         | 11001110               | 0xce
    Case &HCE
        GetMPLength = 1 + 4
        
    'uint 64         | 11001111               | 0xcf
    Case &HCF
        GetMPLength = 1 + 8
        
    'int 8           | 11010000               | 0xd0
    Case &HD0
        GetMPLength = 1 + 1
        
    'int 16          | 11010001               | 0xd1
    Case &HD1
        GetMPLength = 1 + 2
        
    'int 32          | 11010010               | 0xd2
    Case &HD2
        GetMPLength = 1 + 4
        
    'int 64          | 11010011               | 0xd3
    Case &HD3
        GetMPLength = 1 + 8
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetMPLength = 1 + 1 + 1
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetMPLength = 1 + 1 + 2
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetMPLength = 1 + 1 + 4
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetMPLength = 1 + 1 + 8
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetMPLength = 1 + 1 + 16
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        ItemLength = MPBytes(Index + 1)
        GetMPLength = 1 + 1 + ItemLength
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        ItemLength = GetUInt16FromBytes(MPBytes, Index + 1, True)
        GetMPLength = 1 + 2 + ItemLength
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        ItemLength = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
        GetMPLength = 1 + 4 + ItemLength
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        ItemCount = GetUInt16FromBytes(MPBytes, Index + 1, True)
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount, MPBytes, Index + 1 + 2)
        GetMPLength = 1 + 2 + ItemLength
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        ItemCount = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount, MPBytes, Index + 1 + 4)
        GetMPLength = 1 + 4 + ItemLength
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        ItemCount = GetUInt16FromBytes(MPBytes, Index + 1, True)
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount * 2, MPBytes, Index + 1 + 2)
        GetMPLength = 1 + 2 + ItemLength
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        ItemCount = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
        ItemLength = _
            GetMPLengthFromItemMPBytes(ItemCount * 2, MPBytes, Index + 1 + 4)
        GetMPLength = 1 + 4 + ItemLength
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        GetMPLength = 1
        
    End Select
End Function

Private Function GetMPLengthFromItemMPBytes(ByVal ItemCount As Long, _
    MPBytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        Length = Length + GetMPLength(MPBytes, Index + Length)
    Next
    
    GetMPLengthFromItemMPBytes = Length
End Function

Public Function IsMPObject( _
    MPBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case MPBytes(Index)
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    'map 16          | 11011110               | 0xde
    'map 32          | 11011111               | 0xdf
    Case &H80 To &H8F, &HDE, &HDF
        IsMPObject = True
        
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    'array 16        | 11011100               | 0xdc
    'array 32        | 11011101               | 0xdd
    Case &H90 To &H9F, &HDC, &HDD
        #If USE_COLLECTION Then
        IsMPObject = True
        #Else
        IsMPObject = False
        #End If
        
    Case Else
        IsMPObject = False
        
    End Select
End Function

Public Function GetValue(MPBytes() As Byte, Optional Index As Long) As Variant
    Select Case MPBytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        GetValue = GetPositiveFixIntFromMPBytes(MPBytes, Index)
        
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        Set GetValue = GetFixMapFromMPBytes(MPBytes, Index)
        
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        #If USE_COLLECTION Then
        Set GetValue = GetFixArrayFromMPBytes(MPBytes, Index)
        #Else
        GetValue = GetFixArrayFromMPBytes(MPBytes, Index)
        #End If
        
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        GetValue = GetFixStrFromMPBytes(MPBytes, Index)
        
    'nil             | 11000000               | 0xc0
    Case &HC0
        GetValue = GetNilFromMPBytes(MPBytes, Index)
        
    '(never used)    | 11000001               | 0xc1
    Case &HC1
        GetValue = GetNeverUsedFromMPBytes(MPBytes, Index)
        
    'false           | 11000010               | 0xc2
    Case &HC2
        GetValue = GetFalseFromMPBytes(MPBytes, Index)
        
    'true            | 11000011               | 0xc3
    Case &HC3
        GetValue = GetTrueFromMPBytes(MPBytes, Index)
        
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        GetValue = GetBin8FromMPBytes(MPBytes, Index)
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        GetValue = GetBin16FromMPBytes(MPBytes, Index)
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        GetValue = GetBin32FromMPBytes(MPBytes, Index)
        
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetValue = GetExt8FromMPBytes(MPBytes, Index)
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        GetValue = GetExt16FromMPBytes(MPBytes, Index)
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        GetValue = GetExt32FromMPBytes(MPBytes, Index)
        
    'float 32        | 11001010               | 0xca
    Case &HCA
        GetValue = GetFloat32FromMPBytes(MPBytes, Index)
        
    'float 64        | 11001011               | 0xcb
    Case &HCB
        GetValue = GetFloat64FromMPBytes(MPBytes, Index)
        
    'uint 8          | 11001100               | 0xcc
    Case &HCC
        GetValue = GetUInt8FromMPBytes(MPBytes, Index)
        
    'uint 16         | 11001101               | 0xcd
    Case &HCD
        GetValue = GetUInt16FromMPBytes(MPBytes, Index)
        
    'uint 32         | 11001110               | 0xce
    Case &HCE
        GetValue = GetUInt32FromMPBytes(MPBytes, Index)
        
    'uint 64         | 11001111               | 0xcf
    Case &HCF
        GetValue = GetUInt64FromMPBytes(MPBytes, Index)
        
    'int 8           | 11010000               | 0xd0
    Case &HD0
        GetValue = GetInt8FromMPBytes(MPBytes, Index)
        
    'int 16          | 11010001               | 0xd1
    Case &HD1
        GetValue = GetInt16FromMPBytes(MPBytes, Index)
        
    'int 32          | 11010010               | 0xd2
    Case &HD2
        GetValue = GetInt32FromMPBytes(MPBytes, Index)
        
    'int 64          | 11010011               | 0xd3
    Case &HD3
        GetValue = GetInt64FromMPBytes(MPBytes, Index)
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetValue = GetFixExt1FromMPBytes(MPBytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetValue = GetFixExt2FromMPBytes(MPBytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetValue = GetFixExt4FromMPBytes(MPBytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetValue = GetFixExt8FromMPBytes(MPBytes, Index)
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetValue = GetFixExt16FromMPBytes(MPBytes, Index)
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        GetValue = GetStr8FromMPBytes(MPBytes, Index)
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        GetValue = GetStr16FromMPBytes(MPBytes, Index)
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        GetValue = GetStr32FromMPBytes(MPBytes, Index)
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        #If USE_COLLECTION Then
        Set GetValue = GetArray16FromMPBytes(MPBytes, Index)
        #Else
        GetValue = GetArray16FromMPBytes(MPBytes, Index)
        #End If
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        #If USE_COLLECTION Then
        Set GetValue = GetArray32FromMPBytes(MPBytes, Index)
        #Else
        GetValue = GetArray32FromMPBytes(MPBytes, Index)
        #End If
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        Set GetValue = GetMap16FromMPBytes(MPBytes, Index)
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        Set GetValue = GetMap32FromMPBytes(MPBytes, Index)
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        GetValue = GetNegativeFixIntFromMPBytes(MPBytes, Index)
        
    End Select
End Function

'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
'positive fixint stores 7-bit positive integer
'+--------+
'|0XXXXXXX|
'+--------+
'* 0XXXXXXX is 8-bit unsigned integer

Private Function GetPositiveFixIntFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Byte
    
    GetPositiveFixIntFromMPBytes = MPBytes(Index)
End Function

'fixmap          | 1000xxxx               | 0x80 - 0x8f
'fixmap stores a map whose length is upto 15 elements
'+--------+~~~~~~~~~~~~~~~~~+
'|1000XXXX|   N*2 objects   |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetFixMapFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = MPBytes(Index) And &HF
    
    Set GetFixMapFromMPBytes = _
        GetMapFromMPBytes(MPBytes, Index + 1, ItemCount)
End Function

'fixarray        | 1001xxxx               | 0x90 - 0x9f
'fixarray stores an array whose length is upto 15 elements:
'+--------+~~~~~~~~~~~~~~~~~+
'|1001XXXX|    N objects    |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of an array

#If USE_COLLECTION Then

Private Function GetFixArrayFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = MPBytes(Index) And &HF
    
    Set GetFixArrayFromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1, ItemCount)
End Function

#Else

Private Function GetFixArrayFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = MPBytes(Index) And &HF
    
    GetFixArrayFromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1, ItemCount)
End Function

#End If

'fixstr          | 101xxxxx               | 0xa0 - 0xbf
'fixstr stores a byte array whose length is upto 31 bytes:
'+--------+========+
'|101XXXXX|  data  |
'+--------+========+
'* XXXXX is a 5-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetFixStrFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = (MPBytes(Index) And &H1F)
    If Length = 0 Then
        GetFixStrFromMPBytes = ""
        Exit Function
    End If
    
    GetFixStrFromMPBytes = GetStringFromBytes(MPBytes, Index + 1, Length)
End Function

'nil             | 11000000               | 0xc0
'nil:
'+--------+
'|  0xc0  |
'+--------+

Private Function GetNilFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    GetNilFromMPBytes = Null
End Function

'(never used)    | 11000001               | 0xc1

Private Function GetNeverUsedFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    GetNeverUsedFromMPBytes = Empty
End Function

'false           | 11000010               | 0xc2
'false:
'+--------+
'|  0xc2  |
'+--------+

Private Function GetFalseFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Boolean
    
    GetFalseFromMPBytes = False
End Function

'true            | 11000011               | 0xc3
'true:
'+--------+
'|  0xc3  |
'+--------+

Private Function GetTrueFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Boolean
    
    GetTrueFromMPBytes = True
End Function

'bin 8           | 11000100               | 0xc4
'bin 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xc4  |XXXXXXXX|  data  |
'+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is the length of data

Private Function GetBin8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = MPBytes(Index + 1)
    If Length = 0 Then
        GetBin8FromMPBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, MPBytes, Index + 1 + 1, Length
    
    GetBin8FromMPBytes = BinBytes
End Function

'bin 16          | 11000101               | 0xc5
'bin 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xc5  |YYYYYYYY|YYYYYYYY|  data  |
'+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data

Private Function GetBin16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = GetUInt16FromBytes(MPBytes, Index + 1, True)
    If Length = 0 Then
        GetBin16FromMPBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, MPBytes, Index + 1 + 2, Length
    
    GetBin16FromMPBytes = BinBytes
End Function

'bin 32          | 11000110               | 0xc6
'bin 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xc6  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data

Private Function GetBin32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    If Length = 0 Then
        GetBin32FromMPBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, MPBytes, Index + 1 + 4, Length
    
    GetBin32FromMPBytes = BinBytes
End Function

'ext 8           | 11000111               | 0xc7
'ext 8 stores an integer and a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+--------+========+
'|  0xc7  |XXXXXXXX|  type  |  data  |
'+--------+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetExt8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Select Case MPBytes(Index + 1 + 1)
    Case mpExtTimestamp
        GetExt8FromMPBytes = GetExtTime8FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDecimal
        GetExt8FromMPBytes = GetExtDec8FromMPBytes(MPBytes, Index)
        Exit Function
        
    End Select
    
    Dim ExtBytes() As Byte
    
    Dim Length As Byte
    Length = MPBytes(Index + 1)
    If Length = 0 Then
        GetExt8FromMPBytes = ExtBytes
        Exit Function
    End If
    
    ReDim ExtBytes(0 To Length - 1)
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 1 + 1, Length
    
    GetExt8FromMPBytes = ExtBytes
End Function

'ext 16          | 11001000               | 0xc8
'ext 16 stores an integer and a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+--------+========+
'|  0xc8  |YYYYYYYY|YYYYYYYY|  type  |  data  |
'+--------+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetExt16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ExtBytes() As Byte
    
    Dim Length As Long
    Length = GetUInt16FromBytes(MPBytes, Index + 1, True)
    If Length = 0 Then
        GetExt16FromMPBytes = ExtBytes
        Exit Function
    End If
    
    ReDim ExtBytes(0 To Length - 1)
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 2 + 1, Length
    
    GetExt16FromMPBytes = ExtBytes
End Function

'ext 32          | 11001001               | 0xc9
'ext 32 stores an integer and a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+--------+========+
'|  0xc9  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  type  |  data  |
'+--------+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a big-endian 32-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetExt32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ExtBytes() As Byte
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    If Length = 0 Then
        GetExt32FromMPBytes = ExtBytes
        Exit Function
    End If
    
    ReDim ExtBytes(0 To Length - 1)
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 4 + 1, Length
    
    GetExt32FromMPBytes = ExtBytes
End Function

'float 32        | 11001010               | 0xca
'float 32 stores a floating point number in IEEE 754 single precision floating point number format:
'+--------+--------+--------+--------+--------+
'|  0xca  |XXXXXXXX|XXXXXXXX|XXXXXXXX|XXXXXXXX|
'+--------+--------+--------+--------+--------+
'* XXXXXXXX_XXXXXXXX_XXXXXXXX_XXXXXXXX is a big-endian IEEE 754 single precision floating point number.
'  Extension of precision from single-precision to double-precision does not lose precision.

Private Function GetFloat32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Single
    
    GetFloat32FromMPBytes = GetFloat32FromBytes(MPBytes, Index + 1, True)
End Function

'float 64        | 11001011               | 0xcb
'float 64 stores a floating point number in IEEE 754 double precision floating point number format:
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcb  |YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY is a big-endian
'  IEEE 754 double precision floating point number

Private Function GetFloat64FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Double
    
    GetFloat64FromMPBytes = GetFloat64FromBytes(MPBytes, Index + 1, True)
End Function

'uint 8          | 11001100               | 0xcc
'uint 8 stores a 8-bit unsigned integer
'+--------+--------+
'|  0xcc  |ZZZZZZZZ|
'+--------+--------+

Private Function GetUInt8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Byte
    
    GetUInt8FromMPBytes = MPBytes(Index + 1)
End Function

'uint 16         | 11001101               | 0xcd
'uint 16 stores a 16-bit big-endian unsigned integer
'+--------+--------+--------+
'|  0xcd  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+

Private Function GetUInt16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Long
    
    GetUInt16FromMPBytes = GetUInt16FromBytes(MPBytes, Index + 1, True)
End Function
    
'uint 32         | 11001110               | 0xce
'uint 32 stores a 32-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+
'|  0xce  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+

Private Function GetUInt32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    GetUInt32FromMPBytes = GetUInt32FromBytes(MPBytes, Index + 1, True)
End Function

'uint 64         | 11001111               | 0xcf
'uint 64 stores a 64-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+

Private Function GetUInt64FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    GetUInt64FromMPBytes = GetUInt64FromBytes(MPBytes, Index + 1, True)
End Function

'int 8           | 11010000               | 0xd0
'int 8 stores a 8-bit signed integer
'+--------+--------+
'|  0xd0  |ZZZZZZZZ|
'+--------+--------+

Private Function GetInt8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Integer
    
    GetInt8FromMPBytes = GetInt8FromBytes(MPBytes, Index + 1, True)
End Function

'int 16          | 11010001               | 0xd1
'int 16 stores a 16-bit big-endian signed integer
'+--------+--------+--------+
'|  0xd1  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+

Private Function GetInt16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Integer
    
    GetInt16FromMPBytes = GetInt16FromBytes(MPBytes, Index + 1, True)
End Function

'int 32          | 11010010               | 0xd2
'int 32 stores a 32-bit big-endian signed integer
'+--------+--------+--------+--------+--------+
'|  0xd2  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+

Private Function GetInt32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Long
    
    GetInt32FromMPBytes = GetInt32FromBytes(MPBytes, Index + 1, True)
End Function

'int 64          | 11010011               | 0xd3
'int 64 stores a 64-bit big-endian signed integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd3  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+

Private Function GetInt64FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    GetInt64FromMPBytes = GetInt64FromBytes(MPBytes, Index + 1, True)
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExt1FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Select Case MPBytes(Index + 1)
    Case mpExtCurrency
        GetFixExt1FromMPBytes = GetFixExtCur1FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDecimal
        GetFixExt1FromMPBytes = GetFixExtDec1FromMPBytes(MPBytes, Index)
        Exit Function
        
    End Select
    
    Dim ExtBytes(0) As Byte
    ExtBytes(0) = MPBytes(Index + 1 + 1)
    
    GetFixExt1FromMPBytes = ExtBytes
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExt2FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Select Case MPBytes(Index + 1)
    Case mpExtCurrency
        GetFixExt2FromMPBytes = GetFixExtCur2FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDecimal
        GetFixExt2FromMPBytes = GetFixExtDec2FromMPBytes(MPBytes, Index)
        Exit Function
        
    End Select
    
    Dim ExtBytes(0 To 1) As Byte
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 1, 2
    
    GetFixExt2FromMPBytes = ExtBytes
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExt4FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Select Case MPBytes(Index + 1)
    Case mpExtTimestamp
        GetFixExt4FromMPBytes = GetFixExtTime4FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtCurrency
        GetFixExt4FromMPBytes = GetFixExtCur4FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDecimal
        GetFixExt4FromMPBytes = GetFixExtDec4FromMPBytes(MPBytes, Index)
        Exit Function
        
    End Select
    
    Dim ExtBytes(0 To 3) As Byte
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 1, 4
    
    GetFixExt4FromMPBytes = ExtBytes
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExt8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Select Case MPBytes(Index + 1)
    Case mpExtTimestamp
        GetFixExt8FromMPBytes = GetFixExtTime8FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtCurrency
        GetFixExt8FromMPBytes = GetFixExtCur8FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDate
        GetFixExt8FromMPBytes = GetFixExtDate8FromMPBytes(MPBytes, Index)
        Exit Function
        
    Case mpExtDecimal
        GetFixExt8FromMPBytes = GetFixExtDec8FromMPBytes(MPBytes, Index)
        Exit Function
        
    End Select
    
    Dim ExtBytes(0 To 7) As Byte
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 1, 8
    
    GetFixExt8FromMPBytes = ExtBytes
End Function

'fixext 16       | 11011000               | 0xd8
'fixext 16 stores an integer and a byte array whose length is 16 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd8  |  type  |                                  data
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                              data (cont.)                              |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExt16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ExtBytes(0 To 15) As Byte
    CopyBytes ExtBytes, 0, MPBytes, Index + 1 + 1, 16
    
    GetFixExt16FromMPBytes = ExtBytes
End Function

'str 8           | 11011001               | 0xd9
'str 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xd9  |YYYYYYYY|  data  |
'+--------+--------+========+
'* YYYYYYYY is a 8-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetStr8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = MPBytes(Index + 1)
    If Length = 0 Then
        GetStr8FromMPBytes = ""
        Exit Function
    End If
    
    GetStr8FromMPBytes = GetStringFromBytes(MPBytes, Index + 1 + 1, Length)
End Function

'str 16          | 11011010               | 0xda
'str 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xda  |ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetStr16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = GetUInt16FromBytes(MPBytes, Index + 1, True)
    If Length = 0 Then
        GetStr16FromMPBytes = ""
        Exit Function
    End If
    
    GetStr16FromMPBytes = GetStringFromBytes(MPBytes, Index + 1 + 2, Length)
End Function

'str 32          | 11011011               | 0xdb
'str 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xdb  |AAAAAAAA|AAAAAAAA|AAAAAAAA|AAAAAAAA|  data  |
'+--------+--------+--------+--------+--------+========+
'* AAAAAAAA_AAAAAAAA_AAAAAAAA_AAAAAAAA is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string

Private Function GetStr32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    If Length = 0 Then
        GetStr32FromMPBytes = ""
        Exit Function
    End If
    
    GetStr32FromMPBytes = GetStringFromBytes(MPBytes, Index + 1 + 4, Length)
End Function

'array 16        | 11011100               | 0xdc
'array 16 stores an array whose length is upto (2^16)-1 elements:
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdc  |YYYYYYYY|YYYYYYYY|    N objects    |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of an array

#If USE_COLLECTION Then

Private Function GetArray16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(MPBytes, Index + 1, True)
    
    Set GetArray16FromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1 + 2, ItemCount)
End Function

#Else

Private Function GetArray16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(MPBytes, Index + 1, True)
    
    GetArray16FromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1 + 2, ItemCount)
End Function

#End If

'array 32        | 11011101               | 0xdd
'array 32 stores an array whose length is upto (2^32)-1 elements:
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdd  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|    N objects    |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of an array

#If USE_COLLECTION Then

Private Function GetArray32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    
    Set GetArray32FromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1 + 4, ItemCount)
End Function

#Else

Private Function GetArray32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    
    GetArray32FromMPBytes = _
        GetArrayFromMPBytes(MPBytes, Index + 1 + 4, ItemCount)
End Function

#End If

'map 16          | 11011110               | 0xde
'map 16 stores a map whose length is upto (2^16)-1 elements
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xde  |YYYYYYYY|YYYYYYYY|   N*2 objects   |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetMap16FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(MPBytes, Index + 1, True)
    
    Set GetMap16FromMPBytes = _
        GetMapFromMPBytes(MPBytes, Index + 1 + 2, ItemCount)
End Function

'map 32          | 11011111               | 0xdf
'map 32 stores a map whose length is upto (2^32)-1 elements
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|   N*2 objects   |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value

Private Function GetMap32FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(MPBytes, Index + 1, True))
    
    Set GetMap32FromMPBytes = _
        GetMapFromMPBytes(MPBytes, Index + 1 + 4, ItemCount)
End Function

'negative fixint | 111xxxxx               | 0xe0 - 0xff
'negative fixint stores 5-bit negative integer
'+--------+
'|111YYYYY|
'+--------+
'* 111YYYYY is 8-bit signed integer

Private Function GetNegativeFixIntFromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Integer
    
    GetNegativeFixIntFromMPBytes = GetInt8FromBytes(MPBytes, Index, True)
End Function

''
'' MessagePack for VBA - Deserialization - Map Helper
''

Private Function GetMapFromMPBytes( _
    MPBytes() As Byte, Index As Long, ItemCount As Long) As Object
    
    Dim Map As Object
    Set Map = CreateObject("Scripting.Dictionary")
    
    Dim KeyOffset As Long
    Dim ValueOffset As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        ValueOffset = KeyOffset + GetMPLength(MPBytes, Index + KeyOffset)
        
        Map.Add _
            GetValue(MPBytes, Index + KeyOffset), _
            GetValue(MPBytes, Index + ValueOffset)
        
        KeyOffset = ValueOffset + GetMPLength(MPBytes, Index + ValueOffset)
    Next
    
    Set GetMapFromMPBytes = Map
End Function

''
'' MessagePack for VBA - Deserialization - Array Helper
''

#If USE_COLLECTION Then

Private Function GetArrayFromMPBytes( _
    MPBytes() As Byte, Index As Long, ItemCount As Long) As Collection
    
    Dim Collection_ As Collection
    Set Collection_ = New Collection
    
    Dim Offset As Long
    Dim Count As Long
    For Count = 0 To ItemCount - 1
        Collection_.Add GetValue(MPBytes, Index + Offset)
        
        Offset = Offset + GetMPLength(MPBytes, Index + Offset)
    Next
    
    Set GetArrayFromMPBytes = Collection_
End Function

#Else

Private Function GetArrayFromMPBytes( _
    MPBytes() As Byte, Index As Long, ItemCount As Long)
    
    Dim Array_()
    
    If ItemCount = 0 Then
        GetArrayFromMPBytes = Array_
        Exit Function
    End If
    
    ReDim Array_(0 To ItemCount - 1)
    
    Dim Offset As Long
    Dim Count As Long
    For Count = 0 To ItemCount - 1
        If IsMPObject(MPBytes, Index + Offset) Then
            Set Array_(Count) = GetValue(MPBytes, Index + Offset)
        Else
            Array_(Count) = GetValue(MPBytes, Index + Offset)
        End If
        
        Offset = Offset + GetMPLength(MPBytes, Index + Offset)
    Next
    
    GetArrayFromMPBytes = Array_
End Function

#End If

''
'' MessagePack for VBA - Deserialization - Extension
''

'
' -1. Timestamp
'

'ext 8           | 11000111               | 0xc7
'timestamp 96 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 64-bit signed integer and 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+--------+
'|  0xc7  |   12   |   -1   |nanoseconds in 32-bit unsigned int |
'+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                    seconds in 64-bit signed int                        |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* Timestamp 96 format can represent a timestamp in [-292277022657-01-27 08:29:52 UTC, 292277026596-12-04 15:30:08.000000000 UTC) range.
'* In timestamp 64 and timestamp 96 formats, nanoseconds must not be larger than 999999999.

Private Function GetExtTime8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Date
    
    Dim Seconds As Double
    Seconds = GetInt64FromBytes(MPBytes, Index + 1 + 1 + 1 + 4, True)
    
    'GetFixExtTime8FromMPBytes = DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (24# * 60 * 60))
    Seconds = Seconds - Days * (24# * 60 * 60)
    
    GetExtTime8FromMPBytes = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

'fixext 4        | 11010110               | 0xd6
'timestamp 32 stores the number of seconds that have elapsed since 1970-01-01 00:00:00 UTC
'in an 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |   -1   |   seconds in 32-bit unsigned int  |
'+--------+--------+--------+--------+--------+--------+
'* Timestamp 32 format can represent a timestamp in [1970-01-01 00:00:00 UTC, 2106-02-07 06:28:16 UTC) range. Nanoseconds part is 0.

Private Function GetFixExtTime4FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Date
    
    Dim Seconds As Double
    Seconds = GetUInt32FromBytes(MPBytes, Index + 1 + 1, True)
    
    'GetFixExtTime4FromMPBytes = DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (24# * 60 * 60))
    Seconds = Seconds - Days * (24# * 60 * 60)
    
    GetFixExtTime4FromMPBytes = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

'fixext 8        | 11010111               | 0xd7
'timestamp 64 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 32-bit unsigned integers:
'+--------+--------+--------+--------+--------+------|-+--------+--------+--------+--------+
'|  0xd7  |   -1   | nanosec. in 30-bit unsigned int |   seconds in 34-bit unsigned int    |
'+--------+--------+--------+--------+--------+------^-+--------+--------+--------+--------+
'* Timestamp 64 format can represent a timestamp in [1970-01-01 00:00:00.000000000 UTC, 2514-05-30 01:53:04.000000000 UTC) range.

Private Function GetFixExtTime8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Date
    
    Dim SecBytes(0 To 7) As Byte
    CopyBytes SecBytes, 0, MPBytes, Index + 1 + 1, 8
    SecBytes(0) = 0
    SecBytes(1) = 0
    SecBytes(2) = 0
    SecBytes(3) = SecBytes(3) And &H3
    
    Dim Seconds As Double
    Seconds = GetUInt64FromBytes(SecBytes, 0, True)
    
    'GetFixExtTime8FromMPBytes = DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (24# * 60 * 60))
    Seconds = Seconds - Days * (24# * 60 * 60)
    
    GetFixExtTime8FromMPBytes = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

'
' 6. Currency
'

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtCur1FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Currency
    
    Dim CurBytesBE(0 To 7) As Byte
    CurBytesBE(7) = MPBytes(Index + 1 + 1)
    
    GetFixExtCur1FromMPBytes = GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtCur2FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Currency
    
    Dim CurBytesBE(0 To 7) As Byte
    CopyBytes CurBytesBE, 6, MPBytes, Index + 1 + 1, 2
    
    GetFixExtCur2FromMPBytes = GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtCur4FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Currency
    
    Dim CurBytesBE(0 To 7) As Byte
    CopyBytes CurBytesBE, 4, MPBytes, Index + 1 + 1, 4
    
    GetFixExtCur4FromMPBytes = GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtCur8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Currency
    
    GetFixExtCur8FromMPBytes = _
        GetCurrencyFromBytes(MPBytes, Index + 1 + 1, True)
End Function

'
' 7. Date
'

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtDate8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Date
    
    GetFixExtDate8FromMPBytes = GetDateFromBytes(MPBytes, Index + 1 + 1, True)
End Function

'
' 14. Decimal
'

'ext 8           | 11000111               | 0xc7
'ext 8 stores an integer and a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+--------+========+
'|  0xc7  |XXXXXXXX|  type  |  data  |
'+--------+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetExtDec8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Variant
    
    Dim Length As Byte
    Length = MPBytes(Index + 1)
    If Length = 0 Then
        GetExtDec8FromMPBytes = CDec(0)
        Exit Function
    End If
    
    Dim DecBytesBE() As Byte
    ReDim DecBytesBE(0 To 13)
    CopyBytes DecBytesBE, 14 - Length, MPBytes, Index + 1 + 1 + 1, Length
    
    GetExtDec8FromMPBytes = GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtDec1FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Variant
    
    Dim DecBytesBE(0 To 13) As Byte
    DecBytesBE(13) = MPBytes(Index + 1 + 1)
    
    GetFixExtDec1FromMPBytes = GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtDec2FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Variant
    
    Dim DecBytesBE(0 To 13) As Byte
    CopyBytes DecBytesBE, 12, MPBytes, Index + 1 + 1, 2
    
    GetFixExtDec2FromMPBytes = GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtDec4FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Variant
    
    Dim DecBytesBE(0 To 13) As Byte
    CopyBytes DecBytesBE, 10, MPBytes, Index + 1 + 1, 4
    
    GetFixExtDec4FromMPBytes = GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information

Private Function GetFixExtDec8FromMPBytes( _
    MPBytes() As Byte, Optional Index As Long) As Variant
    
    Dim DecBytesBE(0 To 13) As Byte
    CopyBytes DecBytesBE, 6, MPBytes, Index + 1 + 1, 8
    
    GetFixExtDec8FromMPBytes = GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

''
'' MessagePack for VBA - Deserialization - Converter
''

'
' CA. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetFloat32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    GetFloat32FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
End Function

'
' CB. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetFloat64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    GetFloat64FromBytes = GetDoubleFromBytes(Bytes, Index, BigEndian)
End Function

'
' CD. UInt16 - a 16-bit unsigned integer
'

Private Function GetUInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim Bytes4(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes4, 2, Bytes, Index, 2
    Else
        CopyBytes Bytes4, 0, Bytes, Index, 2
    End If
    
    GetUInt16FromBytes = GetLongFromBytes(Bytes4, 0, BigEndian)
End Function

'
' CE. UInt32 - a 32-bit unsigned integer
'

#If USE_LONGLONG Then

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim Bytes8(0 To 7) As Byte
    If BigEndian Then
        CopyBytes Bytes8, 4, Bytes, Index, 4
    Else
        CopyBytes Bytes8, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetLongLongFromBytes(Bytes8, 0, BigEndian)
End Function

#Else

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 10, Bytes, Index, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

#End If

'
' CF. UInt64 - a 64-bit unsigned integer
'

Private Function GetUInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 6, Bytes, Index, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 8
    End If
    
    GetUInt64FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

'
' D0. Int8 - a 8-bit signed integer
'

Private Function GetInt8FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim Bytes2LE(0 To 1) As Byte
    Bytes2LE(0) = Bytes(Index)
    If (Bytes(Index) And &H80) = &H80 Then
        Bytes2LE(1) = &HFF
    Else
        Bytes2LE(1) = 0
    End If
    
    GetInt8FromBytes = GetIntegerFromBytes(Bytes2LE)
End Function

'
' D1. Int16 - a 16-bit signed integer
'

Private Function GetInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    GetInt16FromBytes = GetIntegerFromBytes(Bytes, Index, BigEndian)
End Function

'
' D2. Int32 - a 32-bit signed integer
'

Private Function GetInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    GetInt32FromBytes = GetLongFromBytes(Bytes, Index, BigEndian)
End Function

'
' D3. Int64 - a 64-bit signed integer
'

#If USE_LONGLONG Then

Private Function GetInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    GetInt64FromBytes = GetLongLongFromBytes(Bytes, Index, BigEndian)
End Function

#Else

Private Function GetInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Offset As Long
    Dim Value As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        If (Bytes(Index) And &H80) = &H80 Then
            Bytes14(0) = &H80
            For Offset = 0 To 7
                Bytes14(6 + Offset) = Not Bytes(Index + Offset)
            Next
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian) - 1
        Else
            CopyBytes Bytes14, 6, Bytes, Index, 8
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian)
        End If
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        If (Bytes(Index + 7) And &H80) = &H80 Then
            For Offset = 0 To 7
                Bytes14(0 + Offset) = Not Bytes(Index + Offset)
            Next
            Bytes14(13) = &H80
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian) - 1
        Else
            CopyBytes Bytes14, 0, Bytes, Index, 8
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian)
        End If
    End If
    
    GetInt64FromBytes = Value
End Function

#End If

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetIntegerFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim B2 As Bytes2T
    CopyBytes B2.Bytes, 0, Bytes, Index, 2
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    Dim I As IntegerT
    LSet I = B2
    
    GetIntegerFromBytes = I.Value
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim L As LongT
    LSet L = B4
    
    GetLongFromBytes = L.Value
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetSingleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim S As SingleT
    LSet S = B4
    
    GetSingleFromBytes = S.Value
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetDoubleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DoubleT
    LSet D = B8
    
    GetDoubleFromBytes = D.Value
End Function

'
' 6. Currency - a 64-bit number
'in an integer format, scaled by 10,000 to give a fixed-point number
'with 15 digits to the left of the decimal point and 4 digits to the right.
'

Private Function GetCurrencyFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Currency
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim C As CurrencyT
    LSet C = B8
    
    GetCurrencyFromBytes = C.Value
End Function

'
' 7. Date - an IEEE 754 double precision floating point number
'that represent dates ranging from 1 January 100, to 31 December 9999,
'and times from 0:00:00 to 23:59:59.
'
'When other numeric types are converted to Date,
'values to the left of the decimal represent date information,
'while values to the right of the decimal represent time.
'Midnight is 0 and midday is 0.5.
'

Private Function GetDateFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Date
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DateT
    LSet D = B8
    
    GetDateFromBytes = D.Value
End Function

'
' 8. String
'

Private Function GetStringFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional ByVal Length As Long, _
    Optional Charset As String = "utf-8") As String
    
    If Length = 0 Then
        Length = UBound(Bytes) - Index + 1
    End If
    
    Dim Bytes_() As Byte
    ReDim Bytes_(0 To Length - 1)
    CopyBytes Bytes_, 0, Bytes, Index, Length
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 1 'ADODB.adTypeBinary
        .Write Bytes_
        
        .Position = 0
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        GetStringFromBytes = .ReadText
        
        .Close
    End With
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetDecimalFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    ' BytesXX = Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim BytesXX(0 To 13) As Byte
    CopyBytes BytesXX, 0, Bytes, Index, 14
    
    If BigEndian Then
        ReverseBytes BytesXX
    End If
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    
    ' scale
    BytesX(0) = BytesXX(12)
    
    ' sign
    BytesX(1) = BytesXX(13)
    
    ' data high bytes
    CopyBytes BytesX, 2, BytesXX, 8, 4
    
    ' data low bytes
    CopyBytes BytesX, 6, BytesXX, 0, 8
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    BytesRaw(0) = 14
    BytesRaw(1) = 0
    CopyBytes BytesRaw, 2, BytesX, Index, 14
    
    Dim Value As Variant
    CopyMemory ByVal VarPtr(Value), ByVal VarPtr(BytesRaw(0)), 16
    
    GetDecimalFromBytes = Value
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetLongLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim LL As LongLongT
    LSet LL = B8
    
    GetLongLongFromBytes = LL.Value
End Function

#End If
