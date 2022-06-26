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
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      Caption         =   "Value of Y"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      Caption         =   "Value of X"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Example to evaluate y= 3-7X+2X^2
Option Explicit
Function Poly(x As Double) As Double ' x is what is returned
    Dim dblY
    dblY = 3 - 7 * x + 2 * x ^ 2
    Poly = dblY ' returning the value of the polynomial
End Function
Private Sub txtX_Change()
    txtY = Poly(Val(txtX.Text))
End Sub
