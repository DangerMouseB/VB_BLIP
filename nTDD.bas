Attribute VB_Name = "nTDD"
Option Explicit

Sub runTests()
    runTest1
End Sub

Sub runTest1()
    Dim var As Variant, da() As Double, expectedLength As Long, length As Long, retval As Long, buffer() As Byte, var2 As Variant
    ReDim var(1 To 3) As Variant
    ReDim da(1 To 5, 1 To 5) As Double
    var(1) = "hello"
    var(2) = Empty
    var(3) = da
    expectedLength = DBLengthOfVariantAsBytes(var)
    retval = nCInterface.uBufferSizeForVariant(1, var, length)
    Debug.Assert expectedLength = length
    ReDim buffer(1 To length) As Byte
    retval = nCInterface.uVariantToBuffer(1, var, VarPtr(buffer(1)), length)
    Debug.Assert retval = 0
    ReDim buffer2(1 To length) As Byte
    retval = nCInterface.uVariantFromBuffer(1, VarPtr(buffer(1)), length, var2)
    Debug.Assert UBound(var2) = 3
    Debug.Assert var2(1) = "hello"
    Debug.Assert varType(var2(3)) = vbArray + vbDouble
End Sub
