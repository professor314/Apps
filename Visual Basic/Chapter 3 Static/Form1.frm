VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOne 
      Caption         =   "Add 1 to Box"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' example of a variable
' Douglas Nielson 1/21/03
Option Explicit
Const dblPI As Double = 3.14159

Private Sub cmdOne_Click()
    Static intCount As Integer
    ' add one and display
    intCount = intCount + 1
    txtOutput.Text = intCount
    
End Sub

Private Sub Form_Load()
    Show   ' make form visible
    txtOutput.Text = 2 * dblPI
End Sub

