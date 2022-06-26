VERSION 5.00
Begin VB.Form frmFunction 
   Caption         =   "Function Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtY 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblY 
      Caption         =   "Value of Y"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblX 
      Caption         =   "Value of X"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' example of program to evaluate y = 3 -7x +2x^2
' douglas nielson
Option Explicit
Function Poly(x As Double) As Double ' last is what is returned
    Dim dblY
    dblY = 3 - 7 * x + 2 * x ^ 2
    Poly = dblY ' returning the value of the polynomial
End Function

Private Sub txtX_Change()
    ' call function Poly to get answer
    txtY = Poly(Val(txtX.Text))
End Sub
