Attribute VB_Name = "nCInterface"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit

Private Const BLIP_SCHEMA_1 As Long = 1

Private Const RETVAL_SUCCESS As Long = &H0
Private Const RETVAL_OUT_OF_BUFFER As Long = &H1
Private Const RETVAL_UNKNOWN_SCHEMA As Long = &H2
Private Const RETVAL_UNKNOWN_ERROR = &H3
Private Const RETVAL_SERIALISE_VARIANT_ERROR = &H4


Private myLastErrorState() As Variant

Function uVariantFromBuffer(ByVal schemaID As Long, ByVal pBuffer As Long, ByVal length As Long, oVar As Variant) As Long
    Dim map() As Byte, h As Long
    If Not g_IsInitialised Then initDLL
    
    uVariantFromBuffer = RETVAL_UNKNOWN_ERROR
    
    Select Case schemaID
        
        Case BLIP_SCHEMA_1
            h = createByteArrayMap(map, pBuffer, 1, 1, length, 0, 0, 0, 0).HRESULT
            If h <> S_OK Then Exit Function
            On Error Resume Next
            oVar = DBBytesAsVariant(map, length + 1, 1)
            myLastErrorState = DBErrors_errorState()
            h = releaseArrayMap(map).HRESULT
            If DBErrors_errorStateNumber(myLastErrorState) <> 0 Then uVariantFromBuffer = RETVAL_SERIALISE_VARIANT_ERROR: Exit Function
            If h <> S_OK Then Exit Function
            uVariantFromBuffer = RETVAL_SUCCESS
            
        Case Else
            uVariantFromBuffer = RETVAL_UNKNOWN_SCHEMA
            
    End Select
    
End Function


Function uVariantToBuffer(ByVal schemaID As Long, var As Variant, ByVal pBuffer As Long, ByVal length As Long) As Long
    Dim map() As Byte, h As Long
    If Not g_IsInitialised Then initDLL
    
    uVariantToBuffer = RETVAL_UNKNOWN_ERROR
    
    Select Case schemaID
        
        Case BLIP_SCHEMA_1
            h = createByteArrayMap(map, pBuffer, 1, 1, length, 0, 0, 0, 0).HRESULT
            If h <> S_OK Then Exit Function
            On Error Resume Next
            DBVariantAsBytes var, map, length + 1, 1
            myLastErrorState = DBErrors_errorState()
            h = releaseArrayMap(map).HRESULT
            If DBErrors_errorStateNumber(myLastErrorState) <> 0 Then uVariantToBuffer = RETVAL_SERIALISE_VARIANT_ERROR: Exit Function
            If h <> S_OK Then Exit Function
            uVariantToBuffer = RETVAL_SUCCESS
            
        Case Else
            uVariantToBuffer = RETVAL_UNKNOWN_SCHEMA
            
    End Select
    
End Function


Function uBufferSizeForVariant(ByVal schemaID As Long, var As Variant, oLength As Long) As Long
    If Not g_IsInitialised Then initDLL

    uBufferSizeForVariant = RETVAL_UNKNOWN_ERROR
    
    Select Case schemaID
        
        Case BLIP_SCHEMA_1
            On Error Resume Next
            oLength = DBLengthOfVariantAsBytes(var)
            myLastErrorState = DBErrors_errorState()
            If DBErrors_errorStateNumber(myLastErrorState) <> 0 Then uBufferSizeForVariant = RETVAL_SERIALISE_VARIANT_ERROR: Exit Function
            uBufferSizeForVariant = RETVAL_SUCCESS
            
        Case Else
            uBufferSizeForVariant = RETVAL_UNKNOWN_SCHEMA
            
    End Select
    
End Function
