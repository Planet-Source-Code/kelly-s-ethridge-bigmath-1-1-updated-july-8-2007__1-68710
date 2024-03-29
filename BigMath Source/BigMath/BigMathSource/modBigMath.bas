Attribute VB_Name = "modBigMath"
'    CopyRight (c) 2007 Kelly Ethridge
'
'    This file is part of BigNumberTests.
'
'    BigNumberTests is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    BigNumberTests is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modBigMath
'

''
' This module contains the primary arithmetic algorithms used by the library.
'
' There are two sets of each function in this module. The standard functions are
' optimized and should only be run when compiled (with Integer Overflow turned off.)
'
' The second set of functions are to provide safe versions that can execute within an IDE environment.
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
' This contains all the information about a number. The information can be easily
' passed around as a group instead of trying to pass individual parameters.
'
Public Type Number
    Digits()    As Integer
    Precision   As Long
    Sign        As BigMath.Sign
End Type



''
' This is the basic implementation of a gradeschool style
' addition of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.A
'
Public Function GradeSchoolAdd(ByRef u As Number, ByRef v As Number) As Integer()

#If IDE_SAFE Then
    GradeSchoolAdd = GradeSchoolAddIDE(u, v)
    Exit Function
#End If

    Dim uExtDigit As Long
    Dim vExtDigit As Long

    If u.Sign = Negative Then uExtDigit = &HFFFF&
    If v.Sign = Negative Then vExtDigit = &HFFFF&

    Dim sum() As Integer
    If u.Precision >= v.Precision Then
        ReDim sum(0 To u.Precision)
    Else
        ReDim sum(0 To v.Precision)
    End If

    Dim i       As Long
    Dim k       As Long
    Dim uDigit  As Long
    Dim vDigit  As Long
    For i = 0 To UBound(sum)
        If i < u.Precision Then uDigit = u.Digits(i) And &HFFFF& Else uDigit = uExtDigit
        If i < v.Precision Then vDigit = v.Digits(i) And &HFFFF& Else vDigit = vExtDigit
        
        k = uDigit + vDigit + k ' this is the only change from subtraction
        sum(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolAdd = sum
End Function

''
' This is the basic implementation of a gradeschool style
' subtraction of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.S
'
Public Function GradeSchoolSubtract(ByRef u As Number, ByRef v As Number) As Integer()
    
#If IDE_SAFE Then
    GradeSchoolSubtract = GradeSchoolSubtractIDE(u, v)
    Exit Function
#End If

    Dim uExtDigit As Long
    Dim vExtDigit As Long

    If u.Sign = Negative Then uExtDigit = &HFFFF&
    If v.Sign = Negative Then vExtDigit = &HFFFF&

    Dim difference() As Integer
    If u.Precision >= v.Precision Then
        ReDim difference(0 To u.Precision)
    Else
        ReDim difference(0 To v.Precision)
    End If

    Dim i       As Long
    Dim k       As Long
    Dim uDigit  As Long
    Dim vDigit  As Long
    For i = 0 To UBound(difference)
        If i < u.Precision Then uDigit = u.Digits(i) And &HFFFF& Else uDigit = uExtDigit
        If i < v.Precision Then vDigit = v.Digits(i) And &HFFFF& Else vDigit = vExtDigit
        
        k = uDigit - vDigit + k ' this is the only change from addition
        difference(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolSubtract = difference
End Function

''
' This is a straight forward implementation of Knuth's algorithm.
'
' Ref: The Art of Computer Programming 4.3.1.M
'
Public Function GradeSchoolMultiply(ByRef u As Number, ByRef v As Number) As Integer()

#If IDE_SAFE Then
    GradeSchoolMultiply = GradeSchoolMultiplyIDE(u, v)
    Exit Function
#End If

    Dim product() As Integer
    ReDim product(0 To u.Precision + v.Precision)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    For j = 0 To v.Precision - 1
        Dim d As Long
        d = v.Digits(j) And &HFFFF&
        k = 0
        
        For i = 0 To u.Precision - 1
            k = d * (u.Digits(i) And &HFFFF&) + (product(i + j) And &HFFFF&) + k
            product(i + j) = k And &HFFFF&
            k = ((k And &HFFFF0000) \ &H10000) And &HFFFF&
        Next i
        
        product(i + j) = k And &HFFFF&
    Next j
    
    GradeSchoolMultiply = product
End Function

''
' This is an implementation of Knuth's algorithm.
'
' As simple as division would seem to be in the real world, implementing it at such
' a low level has it's own sets of problems. After careful study of Knuth's algorithm
' I finally came up with the following implmentation. The steps in the book are
' marked inline with the code as close as possible.
'
' Ref: The Art of Computer Programming 4.3.1.D
'
Public Function GradeSchoolDivide(ByRef u As Number, ByRef v As Number, ByRef remainder() As Integer, ByVal includeRemainder As Boolean) As Integer()

#If IDE_SAFE Then
    GradeSchoolDivide = GradeSchoolDivideIDE(u, v, remainder, includeRemainder)
    Exit Function
#End If

    Dim n As Long: n = v.Precision
    Dim m As Long: m = u.Precision - n
    
    ' test if the divisor is shorter than the dividend, if so then just
    ' return a 0 quotient and the dividend as the remainder, is needed.
    If m < 0 Then
        If includeRemainder Then
            ReDim remainder(0 To u.Precision)
            Call CopyMemory(remainder(0), u.Digits(0), u.Precision * 2)
        End If
        
        Call SetSAPtr(GradeSchoolDivide, SafeArrayCreateVector(vbInteger, 0, 0))
        Exit Function
    End If
    
    ' ** D1 Start **
    If (u.Precision - 1) = UBound(u.Digits) Then ReDim Preserve u.Digits(0 To u.Precision)
    u.Digits(u.Precision) = 0
    u.Precision = u.Precision + 1

    Dim d As Long
    d = &H10000 \ (1 + (v.Digits(n - 1) And &HFFFF&))
    
    If d > 1 Then
        Call SingleInPlaceMultiply(u, d)
        Call SingleInPlaceMultiply(v, d)
    End If
    ' ** D1 End **
    
    Dim quotient() As Integer
    ReDim quotient(0 To m + 1)
    
    ' this is the Vn-1 digit used repeatedly in step D3.
    Dim vDigit As Long
    vDigit = v.Digits(n - 1) And &HFFFF&
    
    ' this is the Vn-2 digit used repeatedly in step D3.
    Dim vDigit2 As Long
    If n - 2 >= 0 Then vDigit2 = v.Digits(n - 2) And &HFFFF&
    
    Dim qXu() As Integer    ' cache the array to prevent constant allocate/deallocate
    ReDim qXu(0 To n)       ' the array will be reused for multiplication
    
    ' this is an optimistic caching to be used incase
    ' a negative value is encountered. the same value
    ' will always be used regardless, so cache it here.
    Dim q2 As Long
    Dim r2 As Long
    q2 = &H7FFFFFFF \ vDigit
    r2 = &H7FFFFFFF - (q2 * vDigit) + 1
    
    Dim j       As Long
    Dim rHat    As Long
    Dim qHat    As Long
    
    ' ** D2 Start **
    For j = m To 0 Step -1
        Dim wordu As Long
        
        ' ** D3 Start **
        ' since we are shifting left, it is possible that we could turn wordu
        ' into a negative value and will need to deal with it differently later on.
        wordu = ((u.Digits(j + n) And &HFFFF&) * &H10000) Or (u.Digits(j + n - 1) And &HFFFF&)
        
        ' We have to deal with dividing negatives. They need to work like unsigned.
        If wordu And &H80000000 Then
            Dim q1 As Long
            q1 = (wordu And &H7FFFFFFF) \ vDigit
            rHat = (wordu And &H7FFFFFFF) - (q1 * vDigit) + r2
            
            If rHat >= vDigit Then
                q1 = q1 + 1
                rHat = rHat - vDigit
            End If

            qHat = q1 + q2
        Else
            qHat = wordu \ vDigit
            rHat = wordu - qHat * vDigit
        End If
        
        Do
            If qHat < &H10000 Then
                Dim qHatDigits As Long: qHatDigits = (qHat * (v.Digits(n - 2) And &HFFFF&))
                Dim rHatDigits As Long: rHatDigits = (rHat * &H10000) + (u.Digits(j + n - 2) And &HFFFF&)

                If (qHatDigits - &H80000000) <= (rHatDigits - &H80000000) Then Exit Do
            End If
            
            qHat = qHat - 1
            rHat = rHat + vDigit
        Loop While rHat < &H10000
        ' ** D3 End **
        
        ' ** D4 Start **
        Call SinglePlaceMultiply(v.Digits, n, qHat, qXu)
        
        Dim borrowed As Boolean
        borrowed = MultiInPlaceSubtract(u.Digits, j, qXu)
        ' ** D4 End **
        
        ' ** D5 Start **
        If borrowed Then
            ' ** D6 Start **
            qHat = qHat - 1
            Call MultiInPlaceAdd(u.Digits, j, v.Digits)
            ' ** D6 End **
        End If
        ' ** D5 End **
        
        quotient(j) = qHat And &HFFFF&
    Next j
    ' ** D2 End **
    
    ' ** D8 Start **
    If includeRemainder Then
        If d > 1 Then
            remainder = SinglePlaceDivide(u.Digits, n, d)
        Else
            remainder = u.Digits
        End If
    End If
    ' ** D8 End **
    
    GradeSchoolDivide = quotient
End Function

''
' Performs a single in-place division by 10, returning the remainder.
'
' The buffer is modified by this routine.
'
Public Function SingleInPlaceDivideBy10(ByRef n As Number) As Long
#If IDE_SAFE Then
    SingleInPlaceDivideBy10 = SingleInPlaceDivideBy10IDE(n)
    Exit Function
#End If

    Dim r As Long
    Dim i As Long
    Dim f As Boolean
    Dim d As Long

    For i = n.Precision - 1 To 0 Step -1
        r = (r * &H10000) + (n.Digits(i) And &HFFFF&)
        d = r \ 10
        n.Digits(i) = d And &HFFFF&
        r = r - (d * 10)

        If Not f Then
            If n.Digits(i) = 0 Then
                n.Precision = n.Precision - 1
            Else
                f = True
            End If
        End If
    Next i

    SingleInPlaceDivideBy10 = r
End Function

''
' Performs a Two's Complement on the number, effectively negating it.
'
' The number buffer is modified by this routine. It will also reallocate
' the buffer if necessary.
'
Public Sub Negate(ByRef n As Number)
    ' this is to handle situations like FFFF => FFFF0001.
    If n.Sign = Positive Then
        If n.Digits(n.Precision - 1) And &H8000 Then
            If n.Precision > UBound(n.Digits) Then ReDim Preserve n.Digits(0 To n.Precision)
            n.Digits(n.Precision) = 0
            n.Precision = n.Precision + 1
        End If
    End If

#If IDE_SAFE Then
    Call NegateIDE(n)
    Exit Sub
#End If

    Dim k As Long: k = 1
    Dim i As Long
    For i = 0 To n.Precision - 1
        k = k + ((n.Digits(i) Xor &HFFFF) And &HFFFF&)
        n.Digits(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i

    n.Sign = 0 - n.Sign
End Sub

''
' Performs a single in-place multiplication within the original array.
'
' The number buffer is modified by this routine. It is assumed the buffer
' is large enough to handle the larger result.
'
Public Sub SingleInPlaceMultiply(ByRef n As Number, ByVal value As Long)
#If IDE_SAFE Then
    Call SingleInPlaceMultiplyIDE(n, value)
    Exit Sub
#End If

    Dim result  As Long
    Dim i       As Long

    For i = 0 To n.Precision - 1
        result = result + value * (n.Digits(i) And &HFFFF&)
        n.Digits(i) = result And &HFFFF&
        result = ((result And &HFFFF0000) \ &H10000) And &HFFFF&
    Next i

    If result > 0 Then
        n.Precision = n.Precision + 1
        n.Digits(i) = result And &HFFFF&
    End If
End Sub

''
' Performs a single in-place addition within the original array.
'
' The number buffer must be largest enough to handle any overflow.
'
Public Sub SingleInPlaceAdd(ByRef n As Number, ByVal value As Long)
#If IDE_SAFE Then
    Call SingleInPlaceAddIDE(n, value)
    Exit Sub
#End If

    Dim i As Long
    Do While value > 0
        If i >= n.Precision Then n.Precision = n.Precision + 1
        
        value = value + (n.Digits(i) And &HFFFF&)
        n.Digits(i) = value And &HFFFF&
        value = ((value And &HFFFF0000) \ &H10000) And &HFFFF&
        i = i + 1
    Loop
End Sub

Public Function GetInt(ByVal l As Long) As Integer
    If l And &H8000& Then GetInt = &H8000
    GetInt = GetInt Or (l And &H7FFF&)
End Function

''
' This is a support routine for division.
'
Private Sub SinglePlaceMultiply(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, ByRef w() As Integer)

#If IDE_SAFE Then
    Call SinglePlaceMultiplyIDE(u, Length, v, w)
    Exit Sub
#End If

    Dim k As Long
    Dim i As Long
    For i = 0 To Length - 1
        k = k + (v * (u(i) And &HFFFF&))
        w(i) = k And &HFFFF&
        k = ((k And &HFFFF0000) \ &H10000) And &HFFFF&
    Next i

    w(Length) = k And &HFFFF&
End Sub

''
' This is a support routine for division.
'
Private Function MultiInPlaceSubtract(ByRef u() As Integer, ByVal startIndex As Long, ByRef v() As Integer) As Boolean
    
#If IDE_SAFE Then
    MultiInPlaceSubtract = MultiInPlaceSubtractIDE(u, startIndex, v)
    Exit Function
#End If
    
    Dim k As Long
    Dim result As Long
    Dim d As Long
    Dim i As Long
    Dim j As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = startIndex To UBound(u)
        If j <= ubv Then d = v(j) And &HFFFF& Else d = 0
        
        result = result + ((u(i) And &HFFFF&) - d) + k
        
        If result < 0 Then
            result = result + &H10000
            k = -1
        Else
            k = 0
        End If
        
        u(i) = result And &HFFFF&
        result = ((result And &HFFFF0000) \ &H10000) And &HFFFF&
        j = j + 1
    Next i
    
    MultiInPlaceSubtract = k
End Function

''
' Performs an addition between two arrays, placing the result in the first array.
'
Private Sub MultiInPlaceAdd(ByRef u() As Integer, ByVal startIndex As Long, ByRef v() As Integer)

#If IDE_SAFE Then
    Call MultiInPlaceAddIDE(u, startIndex, v)
    Exit Sub
#End If

    Dim result  As Long
    Dim i       As Long
    Dim j As Long
    Dim d As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = startIndex To UBound(u)
        If j <= ubv Then d = v(j) And &HFFFF& Else d = 0
        
        result = result + (u(i) And &HFFFF&) + d
        u(i) = result And &HFFFF&
        result = ((result And &HFFFF0000) \ &H10000) And &HFFFF&
        j = j + 1
    Next i
End Sub

''
' Divides an array by a single digit (16bit) value, returning the quotient and remainder.
'
Public Function SinglePlaceDivide(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef remainder As Long) As Integer()

#If IDE_SAFE Then
    SinglePlaceDivide = SinglePlaceDivideIDE(u, Length, v, remainder)
    Exit Function
#End If

    Dim r       As Long
    Dim q()     As Integer
    ReDim q(0 To Length)
    
    v = v And &HFFFF&
    
    Dim q2 As Long
    Dim r2 As Long
    q2 = &H7FFFFFFF \ v
    r2 = &H7FFFFFFF - (q2 * v) + 1
    
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        r = r * &H10000 + (u(i) And &HFFFF&)
        
        If r And &H80000000 Then
            Dim q1 As Long
            q1 = (r And &H7FFFFFFF) \ v
            r = (r And &H7FFFFFFF) - (q1 * v) + r2

            If r >= v Then
                q1 = q1 + 1
                r = r - v
            End If

            q(i) = q1 + q2
        Else
            q(i) = r \ v
            r = r - (q(i) And &HFFFF&) * v
        End If
    Next i
    
    remainder = r
    SinglePlaceDivide = q
End Function

Public Sub ApplyTwosComplement(ByRef n() As Integer)
#If IDE_SAFE Then
    Call ApplyTwosComplementIDE(n)
    Exit Sub
#End If

    Dim c As Long: c = 1
    Dim i As Long
    For i = 0 To UBound(n)
        c = ((n(i) Xor &HFFFF) And &HFFFF&) + c
        n(i) = c And &HFFFF&
        c = (c And &HFFFF0000) \ &H10000
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   These are IDE safe versions of the math routines.
'
' These are called by the original routines if the IDE_SAFE
' flag is set.
'
' The routines are not optimized, they are provided only to
' allow the library to function safely in an IDE environment.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If IDE_SAFE Then

Private Sub ApplyTwosComplementIDE(ByRef n() As Integer)
    Dim c As Long: c = 1
    Dim i As Long
    For i = 0 To UBound(n)
        c = (GetLong(n(i) Xor &HFFFF)) + c
        n(i) = GetInt(c)
        c = RightShift16(c)
    Next i
End Sub

Private Function GradeSchoolMultiplyIDE(ByRef u As Number, ByRef v As Number) As Integer()
    Dim product() As Integer
    ReDim product(0 To u.Precision + v.Precision)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    For i = 0 To v.Precision - 1
        k = 0
        For j = 0 To u.Precision - 1
            k = UInt16x16To32(v.Digits(i), u.Digits(j)) + GetLong(product(i + j)) + k
            product(i + j) = GetInt(k)
            k = RightShift16(k)
        Next j
        product(i + j) = GetInt(k)
    Next i
    
    GradeSchoolMultiplyIDE = product
End Function

Private Sub SingleInPlaceMultiplyIDE(ByRef n As Number, ByVal value As Long)
    Dim k As Long
    Dim i As Long
    
    For i = 0 To n.Precision - 1
        k = UInt16x16To32(n.Digits(i), value) + k
        n.Digits(i) = GetInt(k)
        k = RightShift16(k)
    Next i
    
    If k Then
        n.Digits(n.Precision) = GetInt(k)
        n.Precision = n.Precision + 1
    End If
End Sub

Private Sub SingleInPlaceAddIDE(ByRef n As Number, ByVal value As Integer)
    Dim i As Long
    Dim k As Long
    k = GetLong(value)
    
    Do While k > 0
        If i >= n.Precision Then n.Precision = n.Precision + 1
        
        k = GetLong(n.Digits(i)) + k
        n.Digits(i) = GetInt(k)
        k = RightShift16(k)
        i = i + 1
    Loop
End Sub

Private Sub NegateIDE(ByRef n As Number)
    Dim k As Long: k = 1
    Dim i As Long
    
    For i = 0 To n.Precision - 1
        k = k + GetLong(n.Digits(i) Xor &HFFFF)
        n.Digits(i) = GetInt(k)
        k = RightShift16(k)
    Next i
    
    n.Sign = 0 - n.Sign
End Sub

Private Function SingleInPlaceDivideBy10IDE(ByRef n As Number) As Long
    Dim r As Long
    Dim i As Long
    Dim f As Boolean
    Dim d As Long
    
    For i = n.Precision - 1 To 0 Step -1
        r = (r * &H10000) + GetLong(n.Digits(i))
        d = r \ 10
        n.Digits(i) = GetInt(d)
        r = r - (d * 10)
        
        If Not f Then
            If n.Digits(i) = 0 Then
                n.Precision = n.Precision - 1
            Else
                f = True
            End If
        End If
    Next i

    SingleInPlaceDivideBy10IDE = r
End Function

Private Function GradeSchoolAddIDE(ByRef u As Number, ByRef v As Number) As Integer()
    Dim uExtDigit As Long
    Dim vExtDigit As Long

    If u.Sign = Negative Then uExtDigit = &HFFFF&
    If v.Sign = Negative Then vExtDigit = &HFFFF&

    Dim sum() As Integer
    If u.Precision >= v.Precision Then
        ReDim sum(0 To u.Precision)
    Else
        ReDim sum(0 To v.Precision)
    End If
    
    Dim i       As Long
    Dim k       As Long
    Dim uDigit  As Long
    Dim vDigit  As Long
    For i = 0 To UBound(sum)
        If i < u.Precision Then uDigit = GetLong(u.Digits(i)) Else uDigit = uExtDigit
        If i < v.Precision Then vDigit = GetLong(v.Digits(i)) Else vDigit = vExtDigit

        k = uDigit + vDigit + k ' this is the only change for the subtraction
        sum(i) = GetInt(k)
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolAddIDE = sum
End Function

Private Function GradeSchoolSubtractIDE(ByRef u As Number, ByRef v As Number) As Integer()
    Dim uExtDigit As Long
    Dim vExtDigit As Long

    If u.Sign = Negative Then uExtDigit = &HFFFF&
    If v.Sign = Negative Then vExtDigit = &HFFFF&

    Dim difference() As Integer
    If u.Precision >= v.Precision Then
        ReDim difference(0 To u.Precision)
    Else
        ReDim difference(0 To v.Precision)
    End If
    
    Dim i       As Long
    Dim k       As Long
    Dim uDigit  As Long
    Dim vDigit  As Long
    For i = 0 To UBound(difference)
        If i < u.Precision Then uDigit = GetLong(u.Digits(i)) Else uDigit = uExtDigit
        If i < v.Precision Then vDigit = GetLong(v.Digits(i)) Else vDigit = vExtDigit

        k = uDigit - vDigit + k ' this is the only change for the addition
        difference(i) = GetInt(k)
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolSubtractIDE = difference
End Function

Private Function GradeSchoolDivideIDE(ByRef u As Number, ByRef v As Number, ByRef remainder() As Integer, ByVal includeRemainder As Boolean) As Integer()
    Dim n As Long: n = v.Precision
    Dim m As Long: m = u.Precision - n
    
    If m < 0 Then
        If includeRemainder Then
            ReDim remainder(0 To u.Precision)
            Call CopyMemory(remainder(0), u.Digits(0), u.Precision * 2)
        End If
        
        Call SetSAPtr(GradeSchoolDivideIDE, SafeArrayCreateVector(vbInteger, 0, 0))
        Exit Function
    End If
    
    If (u.Precision - 1) = UBound(u.Digits) Then ReDim Preserve u.Digits(0 To u.Precision)
    u.Digits(u.Precision) = 0
    u.Precision = u.Precision + 1
        
    Dim d As Long
    d = &H10000 \ (1 + GetLong(v.Digits(n - 1)))
    
    If d > 1 Then
        Call SingleInPlaceMultiply(u, d)
        Call SingleInPlaceMultiply(v, d)
    End If
    
    Dim quotient() As Integer
    ReDim quotient(0 To m + 1)
    
    Dim vDigit As Integer
    vDigit = v.Digits(n - 1)
    
    Dim vDigit2 As Long
    If n - 2 >= 0 Then vDigit2 = GetLong(v.Digits(n - 2))
    
    Dim qTimesu() As Integer
    ReDim qTimesu(0 To n)
    
    Dim j       As Long
    Dim rHat    As Long
    Dim qHat    As Long
    For j = m To 0 Step -1
        Dim wordu As Long
        wordu = Make32(u.Digits(j + n), u.Digits(j + n - 1))
        
        qHat = UInt32d16To32(wordu, vDigit)
        rHat = UInt32m16To32(wordu, vDigit)
        
        Do
            If qHat < &H10000 Then
                If UInt32Compare(UInt32x16To32(qHat, v.Digits(n - 2)), LeftShift16(rHat) + GetLong(u.Digits(j + n - 2))) <= 0 Then
                    Exit Do
                End If
            End If
            
            qHat = qHat - 1
            rHat = rHat + GetLong(vDigit)
        Loop While rHat < &H10000
        
        Call SinglePlaceMultiply(v.Digits, n, qHat, qTimesu)
        
        Dim borrow As Boolean
        borrow = MultiInPlaceSubtract(u.Digits, j, qTimesu)
        
        If borrow Then
            qHat = qHat - 1
            Call MultiInPlaceAdd(u.Digits, j, v.Digits)
        End If
        
        quotient(j) = GetInt(qHat)
    Next j
    
    If includeRemainder Then
        If d > 1 Then
            remainder = SinglePlaceDivide(u.Digits, n, d)
        Else
            remainder = u.Digits
        End If
    End If
    
    GradeSchoolDivideIDE = quotient
End Function

Private Function UInt32x16To32(ByVal x As Long, ByVal y As Integer) As Long
    Dim v As Currency: v = GetLong(y)
    Dim w As Currency: w = (v * x) * 0.0001@
    Call CopyMemory(UInt32x16To32, w, 4)
End Function

Private Function UInt32Compare(ByVal x As Long, ByVal y As Long) As Long
    Dim u As Currency: Call CopyMemory(u, x, 4)
    Dim v As Currency: Call CopyMemory(v, y, 4)
    UInt32Compare = Sgn(u - v)
End Function


Private Sub SinglePlaceMultiplyIDE(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, ByRef w() As Integer)
    Dim k As Long
    Dim i As Long
    
    For i = 0 To Length - 1
        k = k + UInt32x16To32(v, u(i))
        w(i) = GetInt(k)
        k = RightShift16(k)
    Next i

    w(Length) = GetInt(k)
End Sub

Private Function MultiInPlaceSubtractIDE(ByRef u() As Integer, ByVal startIndex As Long, ByRef v() As Integer) As Boolean
    Dim k As Long
    Dim result As Long
    Dim d As Long
    Dim i As Long
    Dim j As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = startIndex To UBound(u)
        If j <= ubv Then d = GetLong(v(j)) Else d = 0
        
        result = result + (GetLong(u(i)) - d) + k
        
        If result < 0 Then
            result = result + &H10000
            k = -1
        Else
            k = 0
        End If
        
        u(i) = GetInt(result)
        result = RightShift16(result)
        j = j + 1
    Next i
    
    MultiInPlaceSubtractIDE = k
End Function

Private Sub MultiInPlaceAddIDE(ByRef u() As Integer, ByVal startIndex As Long, ByRef v() As Integer)
    Dim result  As Long
    Dim i       As Long
    Dim j As Long
    Dim d As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = startIndex To UBound(u)
        If j <= ubv Then d = GetLong(v(j)) Else d = 0
        
        result = result + GetLong(u(i)) + d
        u(i) = GetInt(result)
        result = RightShift16(result)
        j = j + 1
    Next i
End Sub

Private Function SinglePlaceDivideIDE(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef remainder As Long) As Integer()
    Dim q() As Integer
    ReDim q(0 To Length)
    
    Dim r As Long
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        r = r * &H10000 + GetLong(u(i))
        q(i) = GetInt(UInt32d16To32(r, v))
        r = GetInt(UInt32m16To32(r, v))
    Next i
    
    remainder = r
    SinglePlaceDivideIDE = q
End Function

Public Function GetLong(ByVal x As Long) As Long
    GetLong = x And &HFFFF&
End Function

Private Function UInt16x16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim u As Currency: u = GetLong(x)
    Dim v As Currency: v = GetLong(y)
    Dim w As Currency: w = (u * v) * 0.0001@
    Call CopyMemory(UInt16x16To32, w, 4)
End Function

Private Function UInt32d16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim d As Currency
    Call CopyMemory(d, x, 4)
    d = d * 10000@
    UInt32d16To32 = Int(d / GetLong(y))
End Function

Private Function UInt32m16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim q As Currency
    Dim d As Currency
    Dim v As Currency: v = GetLong(y)
    Call CopyMemory(d, x, 4)
    d = d * 10000@
    q = Int(d / v)
    UInt32m16To32 = d - q * v
End Function

Private Function Make32(ByVal x As Integer, ByVal y As Integer) As Long
    Make32 = LeftShift16(GetLong(x)) Or GetLong(y)
End Function

Private Function RightShift16(ByVal x As Long) As Long
    RightShift16 = ((x And &HFFFF0000) \ &H10000) And &HFFFF&
End Function

Private Function LeftShift16(ByVal x As Long) As Long
    If x And &H8000& Then LeftShift16 = &H80000000
    LeftShift16 = LeftShift16 Or ((x And &H7FFF) * &H10000)
End Function

#End If
