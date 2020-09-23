VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtAnswer 
      Height          =   765
      Index           =   2
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   5400
      Width           =   5655
   End
   Begin VB.TextBox txtAnswer 
      Height          =   765
      Index           =   1
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4560
      Width           =   5655
   End
   Begin VB.TextBox txtAnswer 
      Height          =   765
      Index           =   0
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3720
      Width           =   5655
   End
   Begin VB.CommandButton cmdPow 
      Caption         =   "Pow"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "Subtract"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "Multiply"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "Divide"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtOperand2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   5655
   End
   Begin VB.TextBox txtOperand2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   5655
   End
   Begin VB.TextBox txtOperand2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox txtOperand1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox txtOperand1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5655
   End
   Begin VB.TextBox txtOperand1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label Label12 
      Caption         =   "0x"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "0b"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "d"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "0x"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "0b"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "d"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "0x"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "0b"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "d"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Operand 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Operand 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mArg1 As BigInteger
Private mArg2 As BigInteger



Private Sub cmdAdd_Click()
    Call DisplayAnswer(mArg1.Add(mArg2))
End Sub

Private Sub cmdClear_Click()
    Call Clear
End Sub

Private Sub cmdDivide_Click()
    Dim r As BigInteger
    Call DisplayAnswer(mArg1.DivRem(mArg2, r), r)
End Sub

Private Sub cmdMultiply_Click()
    Call DisplayAnswer(mArg1.Multiply(mArg2))
End Sub

Private Sub cmdPow_Click()
    Call DisplayAnswer(mArg1.Pow(mArg2))
End Sub

Private Sub cmdSubtract_Click()
    Call DisplayAnswer(mArg1.Subtract(mArg2))
End Sub

Private Sub Form_Load()
    Call Clear
End Sub

Private Sub txtOperand1_Change(Index As Integer)
    Static changing As Boolean
    
    If changing Then Exit Sub
    changing = True
    Set mArg1 = ParseArgument(txtOperand1, Index)
    Call DisplayArgument(mArg1, txtOperand1, Index)
    changing = False
End Sub

Private Sub txtOperand2_Change(Index As Integer)
    Static changing As Boolean
    
    If changing Then Exit Sub
    changing = True
    Set mArg2 = ParseArgument(txtOperand2, Index)
    DisplayArgument mArg2, txtOperand2, Index
    changing = False
End Sub

Private Sub Clear()
    Set mArg1 = BigInteger.Zero
    Set mArg2 = BigInteger.Zero
    
    Call DisplayArgument(BigInteger.Zero, txtOperand1, -1)
    Call DisplayArgument(BigInteger.Zero, txtOperand2, -1)
    Call DisplayAnswer(BigInteger.Zero)
    
    On Error Resume Next
    txtOperand1(0).SetFocus
End Sub

Private Sub DisplayAnswer(ByVal b As BigInteger, Optional ByVal r As BigInteger)
    If r Is Nothing Then
        txtAnswer(0).Text = b.ToString
        txtAnswer(1).Text = b.ToString("X")
        txtAnswer(2).Text = b.ToString("B")
    Else
        txtAnswer(0).Text = b.ToString & " r: " & r.ToString
        txtAnswer(1).Text = b.ToString("X") & " r: " & r.ToString("X")
        txtAnswer(2).Text = b.ToString("B") & " r: " & r.ToString("B")
    End If
End Sub

Private Function ParseArgument(ByVal boxes As Object, ByVal Index As Long) As BigInteger
    Dim b As BigInteger
    Select Case Index
        Case 0: Call BigInteger.TryParse(boxes(Index), b)
        Case 1: Call BigInteger.TryParse("&h" & boxes(Index), b)
        Case 2: Call BigInteger.TryParse("0b" & boxes(Index), b)
    End Select
    Set ParseArgument = b
End Function

Private Sub DisplayArgument(ByVal b As BigInteger, ByVal boxes As Object, ByVal Index As Long)
    If b Is Nothing Then Set b = BigInteger.Zero
    
    If Index <> 0 Then boxes(0).Text = b.ToString
    If Index <> 1 Then boxes(1).Text = b.ToString("X")
    If Index <> 2 Then boxes(2).Text = b.ToString("B")
End Sub
