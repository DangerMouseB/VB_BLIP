Attribute VB_Name = "nTDD"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011 David Briant - see https://github.com/DangerMouseB
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'************************************************************************************************************************************************
Option Explicit

Sub runTests()
    runTest1
End Sub

Sub runTest1()
    Dim var As Variant, da() As Double, expectedLength As Long, length As Long, retval As Long, buffer() As Byte, var2 As Variant
    ReDim var(1 To 3) As Variant
    ReDim da(1 To 5, 1 To 5) As Double
    da(1, 1) = 1
    var(1) = "hello"
    var(2) = Empty
    var(3) = da
    expectedLength = DBLengthOfVariantAsBytes(var)
    retval = nCInterface.uBufferSizeForVariant(1, var, length)
    Debug.Assert retval = 0
    Debug.Assert expectedLength = length
    ReDim buffer(1 To length) As Byte
    retval = nCInterface.uVariantToBuffer(1, var, VarPtr(buffer(1)), length)
    Debug.Assert retval = 0
    ReDim buffer2(1 To length) As Byte
    retval = nCInterface.uVariantFromBuffer(1, VarPtr(buffer(1)), length, var2, 0)
    Debug.Assert retval = 0
    Debug.Assert UBound(var2) = 3 - 1
    Debug.Assert var2(1 - 1) = "hello"
    Debug.Assert VarType(var2(3 - 1)) = vbArray + vbDouble
    Debug.Assert var2(3 - 1)(1 - 1, 1 - 1) = 1
End Sub

Sub tdd_dict()
    Dim var As Variant, fred As New Dictionary, length As Long, retval As Long, buffer() As Byte, joe As Dictionary
    fred("fred") = 1
    retval = nCInterface.uBufferSizeForVariant(1, fred, length)
    Debug.Assert retval = 0
    ReDim buffer(1 To length) As Byte
    retval = nCInterface.uVariantToBuffer(1, fred, VarPtr(buffer(1)), length)
    Debug.Assert retval = 0
    retval = nCInterface.uVariantFromBuffer(1, VarPtr(buffer(1)), length, var, 1)
    Debug.Assert UBound(var) = 1
    Debug.Assert VarType(var(1)) = vbObject
    Set joe = var(1)
    Debug.Assert joe("fred") = 1
End Sub

Sub testD()
    Dim fred As New Dictionary
    fred(1.2) = 1
    Debug.Assert fred.Keys(0) = 1.2
End Sub

Sub varTest1()
    Dim c As String, d As Variant, buffer() As Byte, buf2() As Byte, ptr1 As Long, ptr2 As Long, ptr3 As Long, ptr4 As Long, ptr5 As Long
    Const UNICODE_NULL_TERMINATOR_LENGTH As Long = 2
    Const SIZE_LENGTH As Long = 4
    c = "Hello"
    d = c
    buffer = getBuffer(d)
    apiCopyMemory ptr1, buffer(9), 4
    apiCopyMemory ptr2, ByVal ptr1, 4
    ptr3 = VarPtr(c)
    apiCopyMemory ptr4, ByVal ptr3, 4
    ptr5 = StrPtr(c)
    ReDim buf2(1 To Len(c) * 2 + UNICODE_NULL_TERMINATOR_LENGTH + SIZE_LENGTH)
    apiCopyMemory buf2(1), ByVal StrPtr(c) - 4, Len(c) * 2 + UNICODE_NULL_TERMINATOR_LENGTH + SIZE_LENGTH
    Debug.Print ptr1, ptr2, ptr3, ptr4, VarPtr(d), StrPtr(d)
    Stop
    
End Sub

Sub varTest2()
    Dim a() As Byte, b(9 To 10) As Integer, c(7 To 8) As Long, d(5 To 6) As Single, e(-1 To 0) As Double, f(1 To 10) As Boolean, g(1 To 10) As Date, h(1 To 10) As Currency, i(1 To 2) As String, j As Variant, length As Long
    Dim x As Long, buffer() As Byte, result As Variant
    ReDim a(1 To 10) As Byte
    
    For x = 1 To 2
        i(x) = "hello" & x
        e(x - 2) = CDbl(x) / 3.14
    Next
    
    j = Array(CByte(1), CInt(2), CLng(3), CSng(4), CDbl(5), True, CDate("7/7/07"), CCur(8), "9", i, e, a, b, c, d)
    j = Array(j, j)
    
    length = DBLengthOfVariantAsBytes(j)
    ReDim buffer(1 To length) As Byte
    DBVariantAsBytes j, buffer, length + 1, 1
    If DBVerifyStructureOfSerialisedVariant(buffer, length + 1, 1) = False Then Stop
    DBBytesAsVariant buffer, length + 1, 1, result, 1
    
    Stop
'    Debug.Print "Byte: " & DBLengthOfVariantAsBytes(CByte(8))
'    Debug.Print "Integer: " & DBLengthOfVariantAsBytes(CInt(8))
'    Debug.Print "Long: " & DBLengthOfVariantAsBytes(CLng(8))
'    Debug.Print "Single: " & DBLengthOfVariantAsBytes(CSng(8))
'    Debug.Print "Double: " & DBLengthOfVariantAsBytes(CDbl(8))
'    Debug.Print "Boolean: " & DBLengthOfVariantAsBytes(True)
'    Debug.Print "Date: " & DBLengthOfVariantAsBytes(CDate(8))
'    Debug.Print "Currency: " & DBLengthOfVariantAsBytes(CCur(8))
'    Debug.Print "String: " & DBLengthOfVariantAsBytes("hello")
'
'    Debug.Print "Byte(): " & DBLengthOfVariantAsBytes(a)
'    Debug.Print "Integer(): " & DBLengthOfVariantAsBytes(b)
'    Debug.Print "Long(): " & DBLengthOfVariantAsBytes(c)
'    Debug.Print "Single(): " & DBLengthOfVariantAsBytes(d)
'    Debug.Print "Double(): " & DBLengthOfVariantAsBytes(e)
'    Debug.Print "Boolean(): " & DBLengthOfVariantAsBytes(f)
'    Debug.Print "Date(): " & DBLengthOfVariantAsBytes(g)
'    Debug.Print "Currency(): " & DBLengthOfVariantAsBytes(h)
'    Debug.Print "String(): " & DBLengthOfVariantAsBytes(i)
'    Debug.Print "Variant(): " & length
'
'    Stop
'    Erase a
'    Debug.Print "Empty Byte(): " & DBLengthOfVariantAsBytes(a)
'    Debug.Print "Variant(Empty Byte()): " & DBLengthOfVariantAsBytes(Array(a))
End Sub

Sub test_slet()
    Dim fred() As Variant, joe As New Dictionary, sally As Variant, sally2 As Variant, pSally As Long
    DBCreateNewArrayOfVariants fred, 1, 1
    slet fred(1), "hello"
    Debug.Assert fred(1) = "hello"
    slet fred(1), joe
    Debug.Assert fred(1) Is joe
    pSally = VarPtr(sally)
    slet sally, "hello"
    Debug.Assert sally = "hello"
    Debug.Assert pSally = VarPtr(sally)
    slet sally, joe
    Debug.Assert pSally = VarPtr(sally)
    Debug.Assert sally Is joe
    slet sally2, sally
    Debug.Assert sally2 Is joe
End Sub

