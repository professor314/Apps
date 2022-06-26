VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "Display Message"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblCount 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' display multiline message box
' show a counter
Private mintCount As Integer ' define module level counter

Private Sub cmdCount_Click()
    mintCount = mintCount + 1       ' add one to counter
    lblCount = FormatNumber(mintCount, 0) ' format display
End Sub

Private Sub cmdMessage_Click()
    ' shows how to have a two line message
    ' Chr$(13) & Chr$(10) is same as vbCrLf
    MsgBox "Name: " & txtName.Text & _
        vbCrLf & _
        "Amount: " & txtAmount.Text, vbOKOnly, _
        "My Message Box"
End Sub
