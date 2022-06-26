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
   Begin VB.CommandButton cmd1 
      Caption         =   "add one to box"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtoutput 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const dblPI As Double = 3.14159

Private Sub cmd1_Click()
    Static intcount As Integer
    ' add one and display
    intcount = intcount + 1
    txtoutput.Text = intcount
End Sub

Private Sub Form_Load()
    Show ' make form visible
    txtoutput.Text = 2 * dblPI
End Sub
    
