Attribute VB_Name = "modMain"
'
' This just demonstrates how to use some of the features available when dealing
' with the BigMath library.
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

Private Sub Main()
    Creating
    Displaying
    DoingMath
End Sub


Public Sub Creating()
    ' An uninstantiated variable will be treated as zero (BigInteger.Zero)
    ' for all methods called in the BigInteger.<method> fashion.
    '
    ' BigInteger.Multiply(SomeValue, Nothing) returns zero
    Dim b As BigInteger
    
    ' a default instantiation of a BigInteger is equal to zero
    ' set b = b.Multiply(SomeValue) returns zero
    Set b = New BigInteger
    Debug.Print b.ToString
    
    ' Longs, Integers and Bytes are treated the same and
    ' are pretty straight forward in their use
    Set b = Big.NewBigInteger(-12345)
    Debug.Print b.ToString
    
    ' Doubles and Singles are treated the same. The fraction
    ' is pulled out of the variable to handle the exponent and
    ' extend the value out as far as it can go.
    ' Since floating point numbers can't represent values
    ' exactly, errors may be introduced into the result.
    
    ' returns a value of 123450000000000000000
    Set b = Big.NewBigInteger(1.2345E+20)
    Debug.Print b.ToString
    
    ' returns a value of 1234499999999999868928
    Set b = Big.NewBigInteger(1.2345E+21)
    Debug.Print b.ToString
    
    ' Currency types are simply shifted right 4 places to remove the
    ' fractional part. The binary portion is then copied out, including
    ' any negative leading sign bits.
    '
    ' returns a value of 1234567890
    Set b = Big.NewBigInteger(1234567890.1234@)
    Debug.Print b.ToString
    
    ' Decimal types are supported by copying the 96bit magnitude portion
    ' and using the sign bit to complete as BigInteger.
    '
    ' returns a value of 1234567890123456789012345
    Set b = Big.NewBigInteger(CDec("1234567890123456789012345"))
    Debug.Print b.ToString
    
    ' Strings can be parsed for extremely large initial values. Three formats
    ' are supported: Decimal, Hex, Binary. The Hex and Binary formats are
    ' denoted in the string with leading indicators. If the string cannot
    ' be parsed, then an error is raised. To prevent this, the TryParse
    ' should be used instead.
    '
    ' All these produce the same value
    Set b = BigInteger.Parse("123456789012345678901234567890")
    Debug.Print b.ToString
    Set b = BigInteger.Parse("0x18EE90FF6C373E0EE4E3F0AD2")
    Debug.Print b.ToString("X")
    Set b = BigInteger.Parse("0b1100011101110100100001111111101101100001101110011111000001110111001001110001111110000101011010010")
    Debug.Print b.ToString("B")
    
    Debug.Print BigInteger.TryParse("invalid", b)   ' returns false with no error
    
    ' The all easy way to create a new BigInteger is to use the general
    ' function BInt. This will accept anything, including BigInteger objects
    ' and Nothing.
    '
    Set b = BInt("123456789012345678901234567890")
    Debug.Print b.ToString
End Sub

Public Sub Displaying()
    Dim b As BigInteger
    Set b = BInt(123456)
    
    ' outputs the value in decimal form
    Debug.Print b.ToString
    
    ' outputs the value in hexidecimal form using uppercase letters
    Debug.Print b.ToString("X")
    
    ' outputs the value in hexidecimal form using lowercase letters
    Debug.Print b.ToString("x")
    
    ' outputs the value in a string of binary notation.
    Debug.Print b.ToString("b") ' or b.ToString("B"). this is not case sensitive.
    
    ' outputs the value with a minimum width of 10 digits
    Debug.Print b.ToString("d10")   ' 0000123456
    
    Debug.Print b.ToString("X10")   ' 000001E240
    
    Debug.Print b.ToString("b10")   ' will display normally since there are more than 10 characters.
End Sub

Public Sub DoingMath()
    Dim x As BigInteger
    Dim y As BigInteger
    Dim z As BigInteger
    
    Set x = BInt("123456")
    Set y = BInt("123")
    
    ' The first methods shown require the x variable to be instantiated.
    ' The y variable can be Nothing in any scenerio.
    
    ' Adds two BigIntegers
    Set z = x.Add(y)
    Set z = BigInteger.Add(x, y)
    
    ' Subtracts two BigIntegers
    Set z = x.Subtract(y)
    Set z = BigInteger.Subtract(x, y)
    
    ' Divides, with remainder
    Dim r As BigInteger
    Set z = x.DivRem(y, r)
    Set z = BigInteger.DivRem(x, y, r)
    
    ' Operations can be concatenated
    Set z = x.Pow(y).Add(y).ShiftRight(2).Divide(BInt("987654321987654321"))
    Debug.Print z.ToString
End Sub









