Attribute VB_Name = "Test_MsgPack"
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
'' MessagePack for VBA - Test
''

' Test Counter
Private m_Test_Count As Long
Private m_Test_Success As Long
Private m_Test_Fail As Long

' Array
#Const USE_COLLECTION = True

Public Sub Test_MsgPack()
    Test_Initialize
    
    Test_MsgPack_PositiveFixInt_TestCases
    Test_MsgPack_FixMap_TestCases
    Test_MsgPack_FixArray_TestCases
    Test_MsgPack_FixStr_TestCases
    Test_MsgPack_Nil_TestCases
    Test_MsgPack_False_TestCases
    Test_MsgPack_True_TestCases
    Test_MsgPack_Bin8_TestCases
    Test_MsgPack_Bin16_TestCases
    Test_MsgPack_Bin32_TestCases
    Test_MsgPack_Ext8_TestCases
    Test_MsgPack_Ext16_TestCases
    Test_MsgPack_Ext32_TestCases
    Test_MsgPack_Float32_TestCases
    Test_MsgPack_Float64_TestCases
    Test_MsgPack_UInt8_TestCases
    Test_MsgPack_UInt16_TestCases
    Test_MsgPack_UInt32_TestCases
    Test_MsgPack_UInt64_TestCases
    Test_MsgPack_Int8_TestCases
    Test_MsgPack_Int16_TestCases
    Test_MsgPack_Int32_TestCases
    Test_MsgPack_Int64_TestCases
    Test_MsgPack_FixExt1_TestCases
    Test_MsgPack_FixExt2_TestCases
    Test_MsgPack_FixExt4_TestCases
    Test_MsgPack_FixExt8_TestCases
    Test_MsgPack_FixExt16_TestCases
    Test_MsgPack_Str8_TestCases
    Test_MsgPack_Str16_TestCases
    Test_MsgPack_Str32_TestCases
    Test_MsgPack_Array16_TestCases
    Test_MsgPack_Array32_TestCases
    Test_MsgPack_Map16_TestCases
    Test_MsgPack_Map32_TestCases
    Test_MsgPack_NegativeFixInt_TestCases
    
    Test_MsgPack_Time_Ext8_TestCases
    Test_MsgPack_Time_FixExt4_TestCases
    Test_MsgPack_Time_FixExt8_TestCases
    
    Test_MsgPack_Cur_FixExt1_TestCases
    Test_MsgPack_Cur_FixExt2_TestCases
    Test_MsgPack_Cur_FixExt4_TestCases
    Test_MsgPack_Cur_FixExt8_TestCases
    
    Test_MsgPack_Date_FixExt8_TestCases
    
    Test_MsgPack_Dec_Ext8_TestCases
    Test_MsgPack_Dec_FixExt1_TestCases
    Test_MsgPack_Dec_FixExt2_TestCases
    Test_MsgPack_Dec_FixExt4_TestCases
    Test_MsgPack_Dec_FixExt8_TestCases
    
    Test_Terminate
End Sub

'
' MessagePack for VBA - Test Cases
'

Private Sub Test_MsgPack_PositiveFixInt_TestCases()
    Debug.Print "Target: PositiveFixInt"
    
    Test_MsgPack_Int_Core "00", &H0
    Test_MsgPack_Int_Core "7F", &H7F
End Sub

Public Sub Test_MsgPack_FixMap_TestCases()
    Debug.Print "Target: FixMap"
    
    Test_MsgPack_Map_Core "80"
    Test_MsgPack_Map_Core "81 A1 61 00"
    Test_MsgPack_Map_Core2 "8F", &HF
End Sub

Public Sub Test_MsgPack_FixArray_TestCases()
    Debug.Print "Target: FixArray"
    
    Test_MsgPack_Array_Core "90"
    Test_MsgPack_Array_Core "91 A1 61"
    Test_MsgPack_Array_Core2 "9F", &HF
End Sub

Private Sub Test_MsgPack_FixStr_TestCases()
    Debug.Print "Target: FixStr"
    
    Test_MsgPack_Str_Core "A0", ""
    Test_MsgPack_Str_Core "A1 61", "a"
    Test_MsgPack_Str_Core "A3 E3 81 82", ChrW(&H3042)
    Test_MsgPack_Str_Core _
        "BF 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77 78 79 7A 41 42 43 44 45", _
        "abcdefghijklmnopqrstuvwxyzABCDE"
End Sub

Private Sub Test_MsgPack_Nil_TestCases()
    Debug.Print "Target: Nil"
    Test_MsgPack_Nil_Core "C0", Null
End Sub

Private Sub Test_MsgPack_False_TestCases()
    Debug.Print "Target: False"
    
    Test_MsgPack_Bool_Core "C2", False
End Sub

Private Sub Test_MsgPack_True_TestCases()
    Debug.Print "Target: True"
    
    Test_MsgPack_Bool_Core "C3", True
End Sub

Private Sub Test_MsgPack_Bin8_TestCases()
    Debug.Print "Target: Bin8"
    
    Test_MsgPack_Bin_Core "C4 00", ""
    Test_MsgPack_Bin_Core2 "C4 01", &H1
    Test_MsgPack_Bin_Core2 "C4 FF", &HFF
End Sub

Private Sub Test_MsgPack_Bin16_TestCases()
    Debug.Print "Target: Bin16"
    
    Test_MsgPack_Bin_Core2 "C5 01 00", &H100
    Test_MsgPack_Bin_Core2 "C5 FF FF", &HFFFF&
End Sub

Private Sub Test_MsgPack_Bin32_TestCases()
    Debug.Print "Target: Bin32"
    
    Test_MsgPack_Bin_Core2 "C6 00 01 00 00", &H10000
End Sub

Public Sub Test_MsgPack_Ext8_TestCases()
    'Debug.Print "Target: Ext8"
    
    'Test_MsgPack_Ext_Core "C7 00 01", &H1, ""
    'Test_MsgPack_Ext_Core2 "C7 03 01", &H3
    'Test_MsgPack_Ext_Core2 "C7 05 01", &H5
    'Test_MsgPack_Ext_Core2 "C7 07 01", &H7
    'Test_MsgPack_Ext_Core2 "C7 09 01", &H9
    'Test_MsgPack_Ext_Core2 "C7 0F 01", &HF
    'Test_MsgPack_Ext_Core2 "C7 11 01", &H11
    'Test_MsgPack_Ext_Core2 "C7 FF 01", &HFF
End Sub

Public Sub Test_MsgPack_Ext16_TestCases()
    'Debug.Print "Target: Ext16"
    
    'Test_MsgPack_Ext_Core2 "C8 01 00 01", &H100
    'Test_MsgPack_Ext_Core2 "C8 FF FF 01", &HFFFF&
End Sub

Public Sub Test_MsgPack_Ext32_TestCases()
    'Debug.Print "Target: Ext32"
    
    'Test_MsgPack_Ext_Core2 "C9 00 01 00 00 01", &H10000
End Sub

Private Sub Test_MsgPack_Float32_TestCases()
    Debug.Print "Target: Float32"
    
    Test_MsgPack_Float_Core "CA 41 46 00 00", 12.375!
    Test_MsgPack_Float_Core "CA 3F 80 00 00", 1!
    Test_MsgPack_Float_Core "CA 3F 00 00 00", 0.5
    Test_MsgPack_Float_Core "CA 3E C0 00 00", 0.375
    Test_MsgPack_Float_Core "CA 3E 80 00 00", 0.25
    Test_MsgPack_Float_Core "CA BF 80 00 00", -1!
    
    ' Positive Zero
    Test_MsgPack_Float_Core "CA 00 00 00 00", 0!
    
    ' Positive SubNormal Minimum
    Test_MsgPack_Float_Core "CA 00 00 00 01", 1.401298E-45
    
    ' Positive SubNormal Maximum
    Test_MsgPack_Float_Core "CA 00 7F FF FF", 1.175494E-38
    
    ' Positive Normal Minimum
    Test_MsgPack_Float_Core "CA 00 80 00 00", 1.175494E-38
    
    ' Positive Normal Maximum
    Test_MsgPack_Float_Core "CA 7F 7F FF FF", 3.402823E+38
    
    ' Positive Infinity
    Test_MsgPack_Float_Core "CA 7F 80 00 00", "inf"
    
    ' Positive NaN
    Test_MsgPack_Float_Core "CA 7F FF FF FF", "nan"
    
    ' Negative Zero
    Test_MsgPack_Float_Core "CA 80 00 00 00", -0!
    
    ' Negative SubNormal Minimum
    Test_MsgPack_Float_Core "CA 80 00 00 01", -1.401298E-45
    
    ' Negative SubNormal Maximum
    Test_MsgPack_Float_Core "CA 80 7F FF FF", -1.175494E-38
    
    ' Negative Normal Minimum
    Test_MsgPack_Float_Core "CA 80 80 00 00", -1.175494E-38
    
    ' Negative Normal Maximum
    Test_MsgPack_Float_Core "CA FF 7F FF FF", -3.402823E+38
    
    ' Negative Infinity
    Test_MsgPack_Float_Core "CA FF 80 00 00", "-inf"
    
    ' Negative NaN
    Test_MsgPack_Float_Core "CA FF FF FF FF", "-nan"
End Sub

Private Sub Test_MsgPack_Float64_TestCases()
    Debug.Print "Target: Float64"
    
    Test_MsgPack_Float_Core "CB 40 28 C0 00 00 00 00 00", 12.375
    Test_MsgPack_Float_Core "CB 3F F0 00 00 00 00 00 00", 1#
    Test_MsgPack_Float_Core "CB 3F E0 00 00 00 00 00 00", 0.5
    Test_MsgPack_Float_Core "CB 3F D8 00 00 00 00 00 00", 0.375
    Test_MsgPack_Float_Core "CB 3F D0 00 00 00 00 00 00", 0.25
    Test_MsgPack_Float_Core "CB 3F B9 99 99 99 99 99 9A", 0.1
    Test_MsgPack_Float_Core "CB 3F D5 55 55 55 55 55 55", 1# / 3#
    Test_MsgPack_Float_Core "CB BF F0 00 00 00 00 00 00", -1#
    
    ' Positive Zero
    Test_MsgPack_Float_Core "CB 00 00 00 00 00 00 00 00", 0#
    
    ' Positive SubNormal Minimum
    Test_MsgPack_Float_Core "CB 00 00 00 00 00 00 00 01", _
        4.94065645841247E-324
    
    ' Positive SubNormal Maximum
    Test_MsgPack_Float_Core "CB 00 0F FF FF FF FF FF FF", _
        2.2250738585072E-308
    
    ' Positive Normal Minimum
    Test_MsgPack_Float_Core "CB 00 10 00 00 00 00 00 00", _
        2.2250738585072E-308
    
    ' Positive Normal Maximum
    Test_MsgPack_Float_Core "CB 7F EF FF FF FF FF FF FF", _
        "1.79769313486232E+308"
    
    ' Positive Infinity
    Test_MsgPack_Float_Core "CB 7F F0 00 00 00 00 00 00", "inf"
    
    ' Positive NaN
    Test_MsgPack_Float_Core "CB 7F FF FF FF FF FF FF FF", "nan"
    
    ' Negative Zero
    Test_MsgPack_Float_Core "CB 80 00 00 00 00 00 00 00", -0#
    
    ' Negative SubNormal Minimum
    Test_MsgPack_Float_Core "CB 80 00 00 00 00 00 00 01", _
        -4.94065645841247E-324
    
    ' Negative SubNormal Maximum
    Test_MsgPack_Float_Core "CB 80 0F FF FF FF FF FF FF", _
        -2.2250738585072E-308
    
    ' Negative Normal Minimum
    Test_MsgPack_Float_Core "CB 80 10 00 00 00 00 00 00", _
        -2.2250738585072E-308
    
    ' Negative Normal Maximum
    Test_MsgPack_Float_Core "CB FF EF FF FF FF FF FF FF", _
        "-1.79769313486232E+308"
    
    ' Negative Infinity
    Test_MsgPack_Float_Core "CB FF F0 00 00 00 00 00 00", "-inf"
    
    ' Negative NaN
    Test_MsgPack_Float_Core "CB FF FF FF FF FF FF FF FF", "-nan"
End Sub

Private Sub Test_MsgPack_UInt8_TestCases()
    Debug.Print "Target: UInt8"
    
    Test_MsgPack_Int_Core "CC 80", &H80
    Test_MsgPack_Int_Core "CC FF", &HFF
End Sub

Private Sub Test_MsgPack_UInt16_TestCases()
    Debug.Print "Target: UInt16"
    
    Test_MsgPack_Int_Core "CD 80 00", &H8000&
    Test_MsgPack_Int_Core "CD FF FF", &HFFFF&
End Sub

Private Sub Test_MsgPack_UInt32_TestCases()
#If Win64 Then
    Debug.Print "Target: UInt32"
    
    Test_MsgPack_Int_Core "CE 80 00 00 00", &H80000000^
    Test_MsgPack_Int_Core "CE FF FF FF FF", &HFFFFFFFF^
#End If
End Sub

Private Sub Test_MsgPack_UInt64_TestCases()
    'Debug.Print "Target: UInt64"
    'Test_MsgPack_Int_Core "CF 80 00 00 00 00 00 00 00", _
        CDec("9223372036854775808")
    'Test_MsgPack_Int_Core "CF FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
End Sub

Private Sub Test_MsgPack_Int8_TestCases()
    Debug.Print "Target: Int8"
    
    Test_MsgPack_Int_Core "D0 DF", -33
    Test_MsgPack_Int_Core "D0 80", -128
End Sub

Private Sub Test_MsgPack_Int16_TestCases()
    Debug.Print "Target: Int16"
    
    Test_MsgPack_Int_Core "D1 01 00", &H100
    Test_MsgPack_Int_Core "D1 7F FF", &H7FFF
    
    Test_MsgPack_Int_Core "D1 FF 7F", -129
    Test_MsgPack_Int_Core "D1 80 00", CInt(-32768)
End Sub

Private Sub Test_MsgPack_Int32_TestCases()
    Debug.Print "Target: Int32"
    
    Test_MsgPack_Int_Core "D2 00 01 00 00", &H10000
    Test_MsgPack_Int_Core "D2 7F FF FF FF", &H7FFFFFFF
    
    Test_MsgPack_Int_Core "D2 FF FF 7F FF", CLng("-32769")
    Test_MsgPack_Int_Core "D2 80 00 00 00", CLng("-2147483648")
End Sub

Private Sub Test_MsgPack_Int64_TestCases()
#If Win64 Then
    Debug.Print "Target: Int64"
    
    Test_MsgPack_Int_Core "D3 00 00 00 01 00 00 00 00", _
        CLngLng("&H100000000")
    Test_MsgPack_Int_Core "D3 7F FF FF FF FF FF FF FF", _
        CLngLng("&H7FFFFFFFFFFFFFFF")
    
    Test_MsgPack_Int_Core "D3 FF FF FF FF 7F FF FF FF", _
        CLngLng("-2147483649")
    Test_MsgPack_Int_Core "D3 80 00 00 00 00 00 00 00", _
        CLngLng("-9223372036854775808")
#End If
End Sub

Public Sub Test_MsgPack_FixExt1_TestCases()
    'Debug.Print "Target: FixExt1"
    
    'Test_MsgPack_Ext_Core "D4 01 00", &H1, "00"
    'Test_MsgPack_Ext_Core "D4 01 01", &H1, "01"
    'Test_MsgPack_Ext_Core "D4 01 FF", &H1, "FF"
End Sub

Public Sub Test_MsgPack_FixExt2_TestCases()
    'Debug.Print "Target: FixExt2"
    
    'Test_MsgPack_Ext_Core "D5 01 00 00", &H1, "00 00"
    'Test_MsgPack_Ext_Core "D5 01 00 01", &H1, "00 01"
    'Test_MsgPack_Ext_Core "D5 01 01 00", &H1, "01 00"
    'Test_MsgPack_Ext_Core "D5 01 FF FF", &H1, "FF FF"
End Sub

Public Sub Test_MsgPack_FixExt4_TestCases()
    'Debug.Print "Target: FixExt4"
    
    'Test_MsgPack_Ext_Core "D6 01 00 00 00 00", &H1, "00 00 00 00"
    'Test_MsgPack_Ext_Core "D6 01 00 00 00 01", &H1, "00 00 00 01"
    'Test_MsgPack_Ext_Core "D6 01 00 00 01 00", &H1, "00 00 01 00"
    'Test_MsgPack_Ext_Core "D6 01 00 01 00 00", &H1, "00 01 00 00"
    'Test_MsgPack_Ext_Core "D6 01 01 00 00 00", &H1, "01 00 00 00"
    'Test_MsgPack_Ext_Core "D6 01 FF FF FF FF", &H1, "FF FF FF FF"
End Sub

Public Sub Test_MsgPack_FixExt8_TestCases()
    Debug.Print "Target: FixExt8"
    
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 00 01", &H1, _
        "00 00 00 00 00 00 00 01"
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 01 00", &H1, _
        "00 00 00 00 00 00 01 00"
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 01 00 00", &H1, _
        "00 00 00 00 00 01 00 00"
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 00 01 00 00 00", &H1, _
        "00 00 00 00 01 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 00 00 00 01 00 00 00 00", &H1, _
        "00 00 00 01 00 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 00 00 01 00 00 00 00 00", &H1, _
        "00 00 01 00 00 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 00 01 00 00 00 00 00 00", &H1, _
        "00 01 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 01 00 00 00 00 00 00 00", &H1, _
        "01 00 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core "D7 01 FF FF FF FF FF FF FF FF", &H1, _
        "FF FF FF FF FF FF FF FF"
End Sub

Public Sub Test_MsgPack_FixExt16_TestCases()
    'Debug.Print "Target: FixExt16"
    
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00"
    'Test_MsgPack_Ext_Core _
        "D8 01 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF", &H1, _
        "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
End Sub

Private Sub Test_MsgPack_Str8_TestCases()
    Debug.Print "Target: Str8"
    
    Test_MsgPack_Str_Core _
        "D9 20 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77 78 79 7A 41 42 43 44 45 46", _
        "abcdefghijklmnopqrstuvwxyzABCDEF"
    Test_MsgPack_Str_Core2 "D9 FF", &HFF
End Sub

Private Sub Test_MsgPack_Str16_TestCases()
    Debug.Print "Target: Str16"
    
    Test_MsgPack_Str_Core2 "DA 01 00", &H100
    Test_MsgPack_Str_Core2 "DA FF FF", &HFFFF&
End Sub

Private Sub Test_MsgPack_Str32_TestCases()
    Debug.Print "Target: Str32"
    
    Test_MsgPack_Str_Core2 "DB 00 01 00 00", &H10000
End Sub

Public Sub Test_MsgPack_Array16_TestCases()
    Debug.Print "Target: Array16"
    
    Test_MsgPack_Array_Core2 "DC 00 10", &H10
    Test_MsgPack_Array_Core2 "DC 01 00", &H100
    'Test_MsgPack_Array_Core2 "DC FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Array32_TestCases()
    'Debug.Print "Target: Array32"
    
    'Test_MsgPack_Array_Core2 "DD 00 01 00 00", &H10000
End Sub

Public Sub Test_MsgPack_Map16_TestCases()
    Debug.Print "Target: Map16"
    
    Test_MsgPack_Map_Core2 "DE 00 10", &H10
    Test_MsgPack_Map_Core2 "DE 01 00", &H100
    'Test_MsgPack_Map_Core2 "DE FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Map32_TestCases()
    Debug.Print "Target: Map32"
    
    'Test_MsgPack_Map_Core2 "DF 00 01 00 00", &H10000
End Sub

Private Sub Test_MsgPack_NegativeFixInt_TestCases()
    Debug.Print "Target: NegativeFixInt"
    
    Test_MsgPack_Int_Core "E0", -32
    Test_MsgPack_Int_Core "FF", -1
End Sub

Public Sub Test_MsgPack_Time_Ext8_TestCases()
    Debug.Print "Target: Timestamp Ext8"
    
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 FF FF FF F2 42 A4 97 80", _
        DateSerial(100, 1, 1)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 FF FF FF FF FF FF FF FF", _
        DateSerial(1969, 12, 31) + TimeSerial(23, 59, 59)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 00 00 00 04 00 00 00 00", _
        DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 00 00 00 3A FF F4 41 7F", _
        DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)
End Sub

Public Sub Test_MsgPack_Time_FixExt4_TestCases()
    Debug.Print "Target: Timestamp FixExt4"
    
    Test_MsgPack_Ext_Time_Core _
        "D6 FF 00 00 00 00", DateSerial(1970, 1, 1)
    Test_MsgPack_Ext_Time_Core _
        "D6 FF 7F FF FF FF", DateSerial(2038, 1, 19) + TimeSerial(3, 14, 7)
    Test_MsgPack_Ext_Time_Core _
        "D6 FF FF FF FF FF", DateSerial(2106, 2, 7) + TimeSerial(6, 28, 15)
End Sub

Public Sub Test_MsgPack_Time_FixExt8_TestCases()
    Debug.Print "Target: Timestamp FixExt8"
    
    Test_MsgPack_Ext_Time_Core _
        "D7 FF 00 00 00 01 00 00 00 00", _
        DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16)
    Test_MsgPack_Ext_Time_Core _
        "D7 FF 00 00 00 03 FF FF FF FF", _
        DateSerial(2514, 5, 30) + TimeSerial(1, 53, 3)
End Sub

Public Sub Test_MsgPack_Cur_FixExt1_TestCases()
    Debug.Print "Target: Currency FixExt1"
    
    Test_MsgPack_Ext_Cur_Core "D4 06 00", 0@
    Test_MsgPack_Ext_Cur_Core "D4 06 01", CCur("0.0001")
    Test_MsgPack_Ext_Cur_Core "D4 06 FF", CCur("0.0255")
End Sub

Public Sub Test_MsgPack_Cur_FixExt2_TestCases()
    Debug.Print "Target: Currency FixExt2"
    
    Test_MsgPack_Ext_Cur_Core "D5 06 01 00", CCur("0.0256")
    Test_MsgPack_Ext_Cur_Core "D5 06 FF FF", CCur("6.5535")
    
    Test_MsgPack_Ext_Cur_Core "D5 06 27 10", CCur("1")
End Sub

Public Sub Test_MsgPack_Cur_FixExt4_TestCases()
    Debug.Print "Target: Currency FixExt4"
    
    Test_MsgPack_Ext_Cur_Core "D6 06 00 01 00 00", CCur("6.5536")
    Test_MsgPack_Ext_Cur_Core "D6 06 FF FF FF FF", CCur("429496.7295")
End Sub

Public Sub Test_MsgPack_Cur_FixExt8_TestCases()
    Debug.Print "Target: Currency FixExt8"
    
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 00 00 00 01 00 00 00 00", CCur("429496.7296")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 7F FF FF FF FF FF FF FF", CCur("922337203685477.5807")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 80 00 00 00 00 00 00 00", CCur("-922337203685477.5808")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 FF FF FF FF FF FF FF FF", CCur("-0.0001")
    
    Test_MsgPack_Ext_Cur_Core "D7 06 FF FF FF FF FF FF D8 F0", CCur("-1")
End Sub

Public Sub Test_MsgPack_Date_FixExt8_TestCases()
    Debug.Print "Target: Date FixExt8"
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 C1 24 10 34 00 00 00 00", DateSerial(100, 1, 1)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 00 00 00 00 00 00 00 00", DateSerial(1899, 12, 30)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 41 46 92 40 80 00 00 00", DateSerial(9999, 12, 31)
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 00 00 00 00 00 00 00 00", TimeSerial(0, 0, 0)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 3F E0 00 00 00 00 00 00", TimeSerial(12, 0, 0)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 3F EF FF E7 BA 37 5F 32", TimeSerial(23, 59, 59)
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 41 46 92 40 FF FF 9E E9", _
        DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)
End Sub

Public Sub Test_MsgPack_Dec_Ext8_TestCases()
    Debug.Print "Target: Decimal Ext8"
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0C 0E 00 00 00 01 00 00 00 00 00 00 00 00", _
        CDec("18446744073709551616")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0C 0E FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("79228162514264337593543950335")
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 00 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("0.0000000000000000000000000001")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 00 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("7.9228162514264337593543950335")
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 00 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("-1")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 00 FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-79228162514264337593543950335")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("-0.0000000000000000000000000001")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-7.9228162514264337593543950335")
End Sub

Public Sub Test_MsgPack_Dec_FixExt1_TestCases()
    Debug.Print "Target: Decimal FixExt1"
    
    Test_MsgPack_Ext_Dec_Core "D4 0E 00", CDec(0)
    Test_MsgPack_Ext_Dec_Core "D4 0E 01", CDec("1")
    Test_MsgPack_Ext_Dec_Core "D4 0E FF", CDec("255")
End Sub

Public Sub Test_MsgPack_Dec_FixExt2_TestCases()
    Debug.Print "Target: Decimal FixExt2"
    
    Test_MsgPack_Ext_Dec_Core "D5 0E 01 00", CDec("256")
    Test_MsgPack_Ext_Dec_Core "D5 0E FF FF", CDec("65535")
End Sub

Public Sub Test_MsgPack_Dec_FixExt4_TestCases()
    Debug.Print "Target: Decimal FixExt4"
    
    Test_MsgPack_Ext_Dec_Core "D6 0E 00 01 00 00", CDec("65536")
    Test_MsgPack_Ext_Dec_Core "D6 0E FF FF FF FF", CDec("4294967295")
End Sub

Public Sub Test_MsgPack_Dec_FixExt8_TestCases()
    Debug.Print "Target: Decimal FixExt8"
    
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 00 00 00 01 00 00 00 00", CDec("4294967296")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 7F FF FF FF FF FF FF FF", CDec("9223372036854775807")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 80 00 00 00 00 00 00 00", CDec("9223372036854775808")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E FF FF FF FF FF FF FF FF", CDec("18446744073709551615")
End Sub

'
' MessagePack for VBA - Test Core
'

Private Sub Test_MsgPack_Int_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Map_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Map_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Public Sub Test_MsgPack_Array_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    #If USE_COLLECTION Then
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = MsgPack.GetValue(Bytes)
    #Else
    Dim ExpectedDummy
    ExpectedDummy = Array(Empty)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(Bytes)
    #End If
    
    DebugPrint_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Array_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Private Sub Test_MsgPack_Str_Core(HexStr As String, ExpectedValue As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim OutputValue As String
    OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Sub Test_MsgPack_Nil_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Nil_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Nil_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Private Sub Test_MsgPack_Bool_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Bool_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Bool_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Private Sub Test_MsgPack_Bin_Core(HexStr As String, ExpectedHexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetBytesFromHexString(ExpectedHexStr)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Bin_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Bin_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Public Sub Test_MsgPack_Ext_Core( _
    HexStr As String, ExtType As Byte, ExpectedHexStr As String)
    
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetBytesFromHexString(ExpectedHexStr)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Ext_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Ext_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Private Sub Test_MsgPack_Float_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Float_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Ext_Time_Core(HexBE As String, ExpectedValue As Date)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Date
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Ext_Time_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Ext_Time_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Ext_Cur_Core( _
    HexBE As String, ExpectedValue As Currency)
    
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Currency
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Ext_Cur_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Ext_Cur_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Ext_Date_Core(HexBE As String, ExpectedValue As Date)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Date
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Ext_Date_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue, False)
    
    DebugPrint_Ext_Date_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Ext_Dec_Core( _
    HexBE As String, ExpectedValue As Variant)
    
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Variant
    OutputValue = MsgPack.GetValue(BytesBE)
    
    DebugPrint_Ext_Dec_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputMPBytesBE() As Byte
    OutputMPBytesBE = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Ext_Dec_GetBytes OutputValue, OutputMPBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Test Core - Map
'

Public Sub Test_MsgPack_Map_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestMapBytes(HeadBytes, ElementCount)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Map_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Private Function GetTestMapBytes( _
    HeadBytes() As Byte, ElementCount As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(0 To HeadLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To ElementCount
        AddBytes TestBytes, MsgPack.GetMPBytes("key-" & CStr(Index))
        AddBytes TestBytes, MsgPack.GetMPBytes("value-" & CStr(Index))
    Next
    
    GetTestMapBytes = TestBytes
End Function

'
' MessagePack for VBA - Test Core - Array
'

Public Sub Test_MsgPack_Array_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestArrayBytes(HeadBytes, ElementCount)
    
    #If USE_COLLECTION Then
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = MsgPack.GetValue(Bytes)
    #Else
    Dim ExpectedDummy
    ExpectedDummy = Array(Empty)
    
    Dim OutputValue
    OutputValue = MsgPack.GetValue(Bytes)
    #End If
    
    DebugPrint_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputMPBytes() As Byte
    OutputMPBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Array_GetBytes OutputValue, OutputMPBytes, Bytes
End Sub

Private Function GetTestArrayBytes( _
    HeadBytes() As Byte, ElementCount As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(0 To HeadLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To ElementCount
        AddBytes TestBytes, MsgPack.GetMPBytes(Index)
    Next
    
    GetTestArrayBytes = TestBytes
End Function

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

'
' MessagePack for VBA - Test Core - String
'

Private Sub Test_MsgPack_Str_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestStrBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue As String
    ExpectedValue = GetTestStr(DataLength)
    
    Dim OutputValue As String
    OutputValue = MsgPack.GetValue(Bytes)
    
    DebugPrint_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestStr(Length As Long) As String
    Dim TestStr As String
    
    Dim Index As Long
    For Index = 1 To Length
        TestStr = TestStr & Hex(Index Mod 16)
    Next
    
    GetTestStr = TestStr
End Function

Private Function GetTestStrBytes( _
    HeadBytes() As Byte, BodyLength As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(HeadLength + BodyLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To BodyLength
        TestBytes(HeadLength + Index - 1) = Asc(Hex(Index Mod 16))
    Next
    
    GetTestStrBytes = TestBytes
End Function

'
' MessagePack for VBA - Test Core - Binary
'

Private Sub Test_MsgPack_Bin_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim MPBytes() As Byte
    MPBytes = GetTestBinBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetTestBinValue(DataLength)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack.GetValue(MPBytes)
    
    DebugPrint_Bin_GetValue MPBytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack.GetMPBytes(OutputValue)
    
    DebugPrint_Bin_GetBytes OutputValue, OutputBytes, MPBytes
End Sub

Private Function GetTestBinValue(Length As Long) As Byte()
    Dim TestValue() As Byte
    ReDim TestValue(0 To Length - 1)
    
    Dim Index As Long
    For Index = 1 To Length
        TestValue(Index - 1) = Index Mod 256
    Next
    
    GetTestBinValue = TestValue
End Function

Private Function GetTestBinBytes( _
    HeadBytes() As Byte, BodyLength As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(HeadLength + BodyLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To BodyLength
        TestBytes(HeadLength + Index - 1) = Index Mod 256
    Next
    
    GetTestBinBytes = TestBytes
End Function

'
' MessagePack for VBA - Test - Debug.Print - Integer
'

Private Sub DebugPrint_Int_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    If VarType(Value) = vbDecimal Then
        DebugPrint_GetMPBytes CStr(Value), OutputMPBytes, ExpectedMPBytes
    Else
        DebugPrint_GetMPBytes _
            CStr(Value) & " (" & Hex(Value) & ")", _
            OutputMPBytes, ExpectedMPBytes
    End If
End Sub

Private Sub DebugPrint_Int_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    If (VarType(OutputValue) = vbDecimal) Or _
        (VarType(ExpectedValue) = vbDecimal) Then
        
        DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue), CStr(ExpectedValue)
    Else
        DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue) & " (" & Hex(OutputValue) & ")", _
            CStr(ExpectedValue) & " (" & Hex(ExpectedValue) & ")"
    End If
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Map
'

Private Sub DebugPrint_Map_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes _
        "(" & TypeName(Value) & ")", OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Map_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    DebugPrint_GetValue MPBytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Array
'

Private Sub DebugPrint_Array_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes _
        "(" & TypeName(Value) & ")", OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Array_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    DebugPrint_GetValue MPBytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

'
' MessagePack for VBA - Test - Debug.Print - String
'

Private Sub DebugPrint_Str_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes Value, OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Str_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        OutputValue, ExpectedValue
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Nil
'

Private Sub DebugPrint_Nil_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes _
        IIf(IsNull(Value), "Null", "not Null"), OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Nil_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue MPBytes, _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null"), _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null")
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Boolean
'

Private Sub DebugPrint_Bool_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes CStr(Value), OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Bool_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Binary
'

Private Sub DebugPrint_Bin_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    Dim HexString As String
    HexString = GetHexStringFromBytes(Value, , , " ")
    
    DebugPrint_GetMPBytes HexString, OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Bin_GetValue( _
    MPBytes() As Byte, OutputValue() As Byte, ExpectedValue() As Byte)
    
    Dim OutputHexString As String
    OutputHexString = GetHexStringFromBytes(OutputValue, , , " ")
    
    Dim ExpectedHexString As String
    ExpectedHexString = GetHexStringFromBytes(ExpectedValue, , , " ")
    
    DebugPrint_GetValue MPBytes, OutputHexString, ExpectedHexString, _
        OutputHexString, ExpectedHexString
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Float
'

Private Sub DebugPrint_Float_GetBytes( _
    Value, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes CStr(Value), OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Float_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue MPBytes, _
        CStr(OutputValue), CStr(ExpectedValue), _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Extension - Timestamp
'

Private Sub DebugPrint_Ext_Time_GetBytes( _
    Value As Date, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes _
        FormatDateTime(Value, vbLongDate) & " " & _
        FormatDateTime(Value, vbLongTime), _
        OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Ext_Time_GetValue( _
    MPBytes() As Byte, OutputValue As Date, ExpectedValue As Date)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        FormatDateTime(OutputValue, vbLongDate) & " " & _
        FormatDateTime(OutputValue, vbLongTime), _
        FormatDateTime(ExpectedValue, vbLongDate) & " " & _
        FormatDateTime(ExpectedValue, vbLongTime)
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Extension - Currency
'

Private Sub DebugPrint_Ext_Cur_GetBytes( _
    Value As Currency, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes CStr(Value), OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Ext_Cur_GetValue( _
    MPBytes() As Byte, OutputValue As Currency, ExpectedValue As Currency)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Extension - Date
'

Private Sub DebugPrint_Ext_Date_GetBytes( _
    Value As Date, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes _
        FormatDateTime(Value, vbLongDate) & " " & _
        FormatDateTime(Value, vbLongTime), _
        OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Ext_Date_GetValue( _
    MPBytes() As Byte, OutputValue As Date, ExpectedValue As Date)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        FormatDateTime(OutputValue, vbLongDate) & " " & _
        FormatDateTime(OutputValue, vbLongTime), _
        FormatDateTime(ExpectedValue, vbLongDate) & " " & _
        FormatDateTime(ExpectedValue, vbLongTime)
End Sub

'
' MessagePack for VBA - Test - Debug.Print - Extension - Decimal
'

Private Sub DebugPrint_Ext_Dec_GetBytes( _
    Value As Variant, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    DebugPrint_GetMPBytes CStr(Value), OutputMPBytes, ExpectedMPBytes
End Sub

Private Sub DebugPrint_Ext_Dec_GetValue( _
    MPBytes() As Byte, OutputValue As Variant, ExpectedValue As Variant)
    
    DebugPrint_GetValue MPBytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

''
'' Message Pack for VBA - Test Counter
''

Private Property Get Test_Count() As Long
    Test_Count = m_Test_Count
End Property

Private Sub Test_Initialize()
    m_Test_Count = 0
    m_Test_Success = 0
    m_Test_Fail = 0
End Sub

Private Sub Test_Countup(bSuccess As Boolean)
    m_Test_Count = m_Test_Count + 1
    If bSuccess Then
        m_Test_Success = m_Test_Success + 1
    Else
        m_Test_Fail = m_Test_Fail + 1
    End If
End Sub

Private Sub Test_Terminate()
    Debug.Print _
        "Count: " & CStr(m_Test_Count) & ", " & _
        "Success: " & CStr(m_Test_Success) & ", " & _
        "Fail: " & CStr(m_Test_Fail)
End Sub

''
'' Message Pack for VBA - Test - Debug.Print
''

Private Sub DebugPrint_GetMPBytes( _
    Source, OutputMPBytes() As Byte, ExpectedMPBytes() As Byte)
    
    Dim bSuccess As Boolean
    bSuccess = CompareBytes(OutputMPBytes, ExpectedMPBytes)
    
    Test_Countup bSuccess
    
    Dim OutputMPBytesStr As String
    OutputMPBytesStr = GetHexStringFromBytes(OutputMPBytes, , , " ")
    
    Dim ExpectedMPBytesStr As String
    ExpectedMPBytesStr = GetHexStringFromBytes(ExpectedMPBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & Source & _
        " Output: " & OutputMPBytesStr & _
        " Expect: " & ExpectedMPBytesStr
End Sub

Private Sub DebugPrint_GetValue( _
    MPBytes() As Byte, OutputValue, ExpectedValue, Output, Expect)
    
    Dim bSuccess As Boolean
    bSuccess = (OutputValue = ExpectedValue)
    
    Test_Countup bSuccess
    
    Dim MPBytesStr As String
    MPBytesStr = GetHexStringFromBytes(MPBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & MPBytesStr & _
        " Output: " & Output & _
        " Expect: " & Expect
End Sub

''
'' Message Pack for VBA - Test - Byte Array Helper
''

Private Function CompareBytes(Bytes1() As Byte, Bytes2() As Byte) As Boolean
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(Bytes1)
    UB1 = UBound(Bytes1)
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(Bytes2)
    UB2 = UBound(Bytes2)
    
    If (UB1 - LB1 + 1) <> (UB2 - LB2 + 1) Then Exit Function
    
    Dim Index As Long
    For Index = 0 To UB1 - LB1
        If Bytes1(LB1 + Index) <> Bytes2(LB2 + Index) Then Exit Function
    Next
    
    CompareBytes = True
End Function

''
'' Message Pack for VBA - Test - Hex String
''

Private Function GetBytesFromHexString(ByVal Value As String) As Byte()
    Dim Value_ As String
    Dim Index As Long
    For Index = 1 To Len(Value)
        Select Case Mid(Value, Index, 1)
        Case "0" To "9", "A" To "F", "a" To "f"
            Value_ = Value_ & Mid(Value, Index, 1)
        End Select
    Next
    
    Dim Length As Long
    Length = Len(Value_) \ 2
    
    Dim Bytes() As Byte
    
    If Length = 0 Then
        GetBytesFromHexString = Bytes
        Exit Function
    End If
    
    ReDim Bytes(0 To Length - 1)
    
    'Dim Index As Long
    For Index = 0 To Length - 1
        Bytes(Index) = CByte("&H" & Mid(Value_, 1 + Index * 2, 2))
    Next
    
    GetBytesFromHexString = Bytes
End Function

'Private Function GetHexStringFromBytes(Bytes() As Byte,
Private Function GetHexStringFromBytes(Bytes, _
    Optional Index As Long, Optional Length As Long, _
    Optional Separator As String) As String
    
    If Length = 0 Then
        On Error Resume Next
        Length = UBound(Bytes) - Index + 1
        On Error GoTo 0
    End If
    If Length = 0 Then
        GetHexStringFromBytes = ""
        Exit Function
    End If
    
    Dim HexString As String
    HexString = Right("0" & Hex(Bytes(Index)), 2)
    
    Dim Offset As Long
    For Offset = 1 To Length - 1
        HexString = _
            HexString & Separator & Right("0" & Hex(Bytes(Index + Offset)), 2)
    Next
    
    GetHexStringFromBytes = HexString
End Function
