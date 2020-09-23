Attribute VB_Name = "modMain"
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
'    Module: modMain
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

Public Declare Function SafeArrayCreateVector Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal fill As Byte)
'Public Declare Sub GetSAPtr Lib "msvbvm60.dll" Alias "GetMem4" (ByRef Arr() As Any, ByRef Ptr As Long)
Public Declare Sub SetSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Arr() As Any, ByVal Ptr As Long)
Public Declare Sub ClearSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Arr() As Any, Optional ByVal Ptr As Long)
Public Declare Sub PutMem4 Lib "msvbvm60.dll" (ByRef Destination As Any, ByVal value As Long)
Public Declare Sub CopySAPtr Lib "msvbvm60.dll" Alias "GetMem4" (ByRef Src() As Any, ByRef Dst() As Any)

Public Type SAFEARRAY1D
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements   As Long
    lBound      As Long
End Type

Public Big          As New Constructors
Public BigInteger   As New BigIntegerStatic
Public PowersOf2()  As Integer


Public Sub Main()
    Call InitPowersOf2
End Sub

Private Sub InitPowersOf2()
    ReDim PowersOf2(0 To 15)
    Dim i As Long
    For i = 0 To 14
        PowersOf2(i) = 2 ^ i
    Next i
    
    PowersOf2(15) = &H8000
End Sub

