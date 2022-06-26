VERSION 5.00
Begin VB.Form frmMaximum 
   Caption         =   "Find Maximum Value"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMaximum 
      Caption         =   "Find Maximum"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtMaximum 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtValue3 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtValue2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtValue1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Maximum Value:"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Value 3"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Value 2"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Value 1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMaximum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program to find the maximum of three values
' Example for chapter 4

Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdMaximum_Click()
    ' add code here to find the maximum value of
    '  txtValue1, txtValue2, and txtValue2 and
    '  place the result in txtMaximum
    
    
    
    
End Sub

Private Sub Form_Load()
    Show
    txtValue1.SetFocus      ' focus to first field
End Sub
