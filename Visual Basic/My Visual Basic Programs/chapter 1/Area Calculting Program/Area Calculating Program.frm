VERSION 5.00
Begin VB.Form frmArea 
   Caption         =   "Finding the area of a circle"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTan 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtpiX 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtcosine 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtsine 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtsquared 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtArea 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "1.0"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblTan 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tan of X Radians:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Cosine of X radians:"
      Height          =   255
      Left            =   -120
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Pi times X:"
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sine of X radians:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "X squared:"
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "by Sean Connollly"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   6360
      Width           =   3900
   End
   Begin VB.Label lblArea 
      Alignment       =   1  'Right Justify
      Caption         =   "Area of a circle with radius X:"
      Height          =   375
      Left            =   -120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Input X:"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   570
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Variable Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Area of a Circle Program
' by Sean Connolly
Option Explicit

Private Sub txtX_Change()
    ' compute area of a circle from text radius
    If txtX.Text <> "" Then
    txtArea.Text = 3.14159265358979 * CDbl(txtX) ^ 2
    txtsine.Text = Sin(txtX)
    txtsquared.Text = (txtX) ^ 2
    txtpiX.Text = (txtX) * 3.14159265358979
    txtcosine.Text = Cos(txtX)
    End If
End Sub
