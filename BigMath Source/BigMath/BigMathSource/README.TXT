--------------
BigMathLib
--------------

- 1.0 Initial Release
- 1.1 Added Factorial function
      Fixed left bit shifting

'****************************************************************************************
'
' IDE_SAFE is defined in the Conditional Compilation Arguments in the project properties.
'
' Compiler Settings
' - IDE_SAFE = 0
' - All Optimizations On
'
'****************************************************************************************

- What is it?
This library contains classes and methods to perform arithmetic on very large signed integer values.

- Why do it?
I was inspired by working with the BigInteger structure in the new .NET VS Orcas edition. I wanted
to see if something similar could be reporduced in VB6 and this is what I've come up with.
Since the BigInteger has been pulled from the beta 1 of the Orcas edition, I have no way of
knowing how the interface has changed. The reason for the BigInteger removal from the new VS
beta is said to be they were going to implement suggestions and such from user feedback. However,
as I was implementing my version and using their's to generate values to ensure mine is working
correctly, I noticed theirs was indeed failing to calculate correctly.

- How is it done?
The values are represented in the same manner as the integer types supported in VB, just on
a much larger scale. VB (because of the x86 architecture) represents numbers in a binary
notation. Also, it represents negative numbers in a Two's Complement value of the positive.
The BigInteger class represents numbers in the same fashion. Usually it wouldn't matter how
the number is represented internally, as long as it calculates and outputs to strings in
an expected manner. However, since the BigInteger class supports bitwise manipulation,
including all of the logical operations as well as bit shifting, it became necessary to
internally represent the number in a binary notation. This also gives the benefit that VB'ers
will be familiar with the bitwise manipulation used as it is basically the same as when
performing the same functions on the smaller integer datatypes. This is just on a much larger scale.

The internal representation was first built using Byte arrays. After many time trials, I was
dissatisfied with the results. Even though it was much faster than other implementations, it was
not was I was hoping for. I then did some studies using Integer arrays. This would give me
twice the throughput per operation if I could only control the sign correctly. I succeeded, but
only in compiled mode. IDE mode uses safe routines that get around signed calculations, and are
far slower. 

Using Integer arrays increased the division by 2x and the multiplication by 3.3x. It has come
down to the slowest parts of the operations usually have to deal with creating and distroying
the objects themselves. I can live with that for now.

- Resources
The bulk of the information needed to implement the math routines was obtained from
the classic set of Donald Knuth's The Art of Computer Programming books, specifically volume 2.

- Future
I imagine I will expand the library with some form of a BigDecimal or BigFloat. As my knowledge
grows, the more I will determine if I want to pursue such things.
A source forge project is being created to help with support and any updates to the project.

I also might attempt a naked version that uses a Type instead of a class to remove the creation 
and destruction of objects. I'm a class guy, so I'm not fond of exposing the raw data for
direct manipulation. I guess I don't trust the user enough ;) Also, I will probably re-implement
this library in .NET classes.

- Me
Kelly

