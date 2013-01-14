Attribute VB_Name = "nGlobals"
'************************************************************************************************************************************************
'
'    Copyright (c) 2011-2012 David Briant - see https://github.com/DangerMouseB
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

Public Const DLL_NAME As String = "BLIP.dll"        ' This value will change for each new project. Be sure to set it appropriately:

' HRESULTS for this project
Public Const E_WRONG_NUMBER_OF_DIMENSIONS As Long = &H800A0200
Public Const E_TOO_MANY_LOCKS As Long = &H800A0201
Public Const E_UNSUPPORT_TYPE_FOR_CHANGE_OF_DIMENSIONS As Long = &H800A0202


