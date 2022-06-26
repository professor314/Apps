VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Transaction"
      Height          =   1575
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   3855
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optServiceFee 
         Caption         =   "Service Fee"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optDeposit 
         Caption         =   "Deposit"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optWithdrawl 
         Caption         =   "Withdrawl"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "C&alculate"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcurBalance As Currency


Private Sub cmdCalculate_Click()
    ' to update balance
    ' check for valid input in text box
    If Len(txtAmount.Text) = 0 Then 'no value entered
        MsgBox "no value entered", vbCritical + vbOKOnly, "Error Message"
    txtAmount.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtAmount.Text) Then
        MsgBox "non numeric value entered", vbCritical + vbOKOnly, "Error Message"
    txtAmount.SetFocus
        Exit Sub
    End If
    If optWithdrawl.Value = True Then
    mcurBalance = mcurBalance - Val(txtAmount.Text)
    ElseIf optDeposit.Value Then
    mcurBalance = mcurBalance + Val(txtAmount.Text)
    
Exit Sub

Private Sub cmdClear_Click()
    optWithdrawl.Value = True
    cmdAmount.Text = ""
    txtAmount.SetFocus
End Sub

Private Sub Form_Load()
    Show
    txtAmount.SetFocus
End Sub
