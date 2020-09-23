VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constructors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2007 Kelly Ethridge
'
'    This file is part of BigNumber.
'
'    BigNumber is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    BigNumber is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Constructors
'

''
' Provides methods to create new BigInteger instances.
'
Option Explicit

'****************************************************************************************
'
' IDE_SAFE is defined in the Conditional Compilation Arguments in the project properties.
'
' Compiler Settings
' - IDE_SAFE = 0
' - All Optimizations On
'
'****************************************************************************************

''
' Represents the sign values of BigInteger values.
'
' @param Negative The value is a negative number.
' @param Zero The value is zero.
' @param Positive The value is a positive number.
'
Public Enum Sign
    Negative = -1
    Zero = 0
    Positive = 1
End Enum

''
' Creates a new BigInteger instance.
'
' @param value The value to intialize a new instance of <b>BitInteger</b> to.
' @return A new BigInteger instance containing the specified value.
' @remarks The follow datatypes are supported:<br>
' vbLong<br>
' vbInteger<br>
' vbByte<br>
' vbDoube<br>
' vbSingle<br>
' vbCurrency<br>
' vbDecimal<br>
' vbByte Or vbArray<br>
'
Public Function NewBigInteger(ByRef value As Variant) As BigInteger
    Set NewBigInteger = New BigInteger
    Call NewBigInteger.Init(value)
End Function

''
' Attempts to convert a value to a valid BigInteger instance.
'
' @param value The value to be converted to BigInteger.
' @return A new BigInteger instance representing the converted value.
' @remarks The follow datatypes are supported:<br>
' vbString<br>
' vbLong<br>
' vbInteger<br>
' vbByte<br>
' vbDoube<br>
' vbSingle<br>
' vbCurrency<br>
' vbDecimal<br>
' vbByte Or vbArray<br>
' BigInteger (returns same instance)<br>
' Nothing (returns BigInteger.Zero)<br>
'
' <pre>
' Dim b As BigInteger
' Set b = BInt(<i>somevalue</i>)
' </pre>
'
Public Function BInt(ByRef value As Variant) As BigInteger
    Select Case VarType(value)
        Case vbObject, vbDataObject
            If value Is Nothing Then
                Set BInt = BigInteger.Zero
            Else
                Set BInt = value
            End If
            
        Case vbString
            Set BInt = BigInteger.Parse(value)
            
        Case Else
            Set BInt = NewBigInteger(value)

    End Select
End Function
