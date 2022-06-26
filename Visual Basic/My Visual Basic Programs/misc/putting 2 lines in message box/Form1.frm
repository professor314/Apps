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
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "Display Message"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblCount 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintcount As Long

Private Sub cmdCount_Click()
    mintcount = mintcount + 1
    lblCount = FormatNumber(mintcount, 0)
End Sub

'display multi-line text box
Private Sub cmdMessage_Click()
    MsgBox "Name: " & txtName.Text & vbCrLf & _
    "Amount: " & txtAmount.Text, vbOKOnly, "My Message Box"
    End Sub
