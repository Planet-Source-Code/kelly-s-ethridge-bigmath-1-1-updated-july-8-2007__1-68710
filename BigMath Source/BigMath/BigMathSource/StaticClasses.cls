VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticClasses"
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
'    Module: StaticClasses
'

''
' Provides access to common object instances.
'
' <h1>Remarks</h1><br>
' This class is marked GlobalMultiuse, so these methods are accessible by directly
' typing in their name.<br>
' <pre>
' Dim b As BigInteger
' Set b = Big.NewBigInteger(15000)
'
' Set b = BigInteger.Parse("12345678901234567890")
' </pre>
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
' Provides a method to access the constructors directly with Intellisense support.
'
' @return The constructors class.
'
Public Function Big() As Constructors
    Set Big = modMain.Big
End Function

''
' Provides a method to access the BigInteger static methods.
'
' @return The BigInteger static methods.
'
Public Function BigInteger() As BigIntegerStatic
    Set BigInteger = modMain.BigInteger
End Function
