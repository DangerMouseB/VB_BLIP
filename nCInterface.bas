Attribute VB_Name = "nCInterface"
'************************************************************************************************************************************************
'
'    Copyright (c) 2011 David Briant - see https://github.com/DangerMouseB
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

Private Const BLIP_SCHEMA_1 As Long = 1

Private Const RETVAL_SUCCESS As Long = &H0
Private Const RETVAL_OUT_OF_BUFFER As Long = &H1
Private Const RETVAL_UNKNOWN_SCHEMA As Long = &H2
Private Const RETVAL_UNKNOWN_ERROR = &H3
Private Const RETVAL_SERIALISE_VARIANT_ERROR = &H4


Private myLastErrorState() As Variant

Function uVariantFromBuffer(ByVal schemaID As Long, ByVal pBuffer As Long, ByVal length As Long, oVar As Variant, base As Long) As Long
    Dim map() As Byte, h As Long
    If Not g_IsInitialised Then initDLL
    
    uVariantFromBuffer = RETVAL_UNKNOWN_ERROR
    
    Select Case schemaID
        
        Case BLIP_SCHEMA_1
            h = createByteArrayMap(map, pBuffer, 1, 1, length, 0, 0, 0, 0).HRESULT
            If h <> S_OK Then Exit Function
            On Error Resume Next
            DBBytesAsVariant map, length + 1, 1, oVar, base
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
