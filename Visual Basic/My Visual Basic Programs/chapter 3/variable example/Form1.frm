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
   Begin VB.TextBox txtoutput 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
Option Explicit
Private Sub Form_Load()
    Const dblpi As Double = 3.14159
    Show
    txtoutput.Text = 2 * dblpi
End Sub
Private Sub txtoutput_change()
    txtoutput.Text = dblpi
End Sub
