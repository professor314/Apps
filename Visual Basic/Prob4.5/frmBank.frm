VERSION 5.00
Begin VB.Form frmBank 
   Caption         =   "Checking Account Balance"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSummary 
      Caption         =   "&Summary"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame fraTransaction 
      Caption         =   "Transaction"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optService 
         Caption         =   "Service"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optDeposit 
         Caption         =   "Deposit"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWithdraw 
         Caption         =   "Withdrawal"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Amount"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblBalance 
      Caption         =   "$0.00"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Balance:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 4.5 (starting with 4.4
' Douglas Nielson
Option Explicit
Private mcurBalance As Currency  ' running balance
Private mintDepositCount As Integer
Private mcurDepositAmount As Currency
Private mintCheckCount As Integer
Private mcurCheckAmount As Currency
Private mcurFeeAmount As Currency

Private Sub cmdCalculate_Click()
    ' calculate button to update balance
    ' check for valid input in text box
    Dim curAmount As Currency ' amount of withdrawal
    If Len(txtAmount.Text) = 0 Then ' no value entered
        MsgBox "No value entered", vbCritical + _
            vbOKOnly, "Error Message"
        txtAmount.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtAmount.Text) Then
        MsgBox "Non numeric value entered", vbCritical + _
            vbOKOnly, "Error Message"
        txtAmount.SetFocus
        Exit Sub
    End If
    If optWithdraw.Value = True Then  ' have a withdrawal
        ' below is change for problem 4.4 page 168
        curAmount = Val(txtAmount.Text) ' get amount
        If mcurBalance - curAmount < 0# Then ' not enough money
            MsgBox "Overdrawn - " & vbCrLf & _
                "You don't have enough money", _
                vbOKOnly + vbCritical, _
                "Error Message"
        Else    ' they have enough money
            mcurBalance = mcurBalance - curAmount
            ' process summary info
            mintCheckCount = mintCheckCount + 1
            mcurCheckAmount = mcurCheckAmount + curAmount
        End If
    ElseIf optDeposit.Value Then ' have a deposit
        mcurBalance = mcurBalance + Val(txtAmount.Text)
        ' process summary info
        mintDepositCount = mintDepositCount + 1 ' add 1
        mcurDepositAmount = mcurDepositAmount + _
            Val(txtAmount.Text) ' add amount
    Else  ' have a service fee
        mcurBalance = mcurBalance - Val(txtAmount.Text)
        ' process summary info
        mcurFeeAmount = mcurFeeAmount + _
            Val(txtAmount.Text) ' fee amount
    End If
    lblBalance.Caption = FormatCurrency(mcurBalance)
End Sub

Private Sub cmdClear_Click() ' clear last transaction
    optWithdraw.Value = True
    txtAmount.Text = ""
    txtAmount.SetFocus
End Sub

Private Sub cmdExit_Click()
    ' end the program
    End
End Sub

Private Sub cmdSummary_Click()
    MsgBox "Deposits: " & vbTab & _
        FormatNumber(mintDepositCount, 0) & vbTab & _
        FormatCurrency(mcurDepositAmount, 2) & _
        vbCrLf & "Checks:" & vbTab & vbTab & _
        FormatNumber(mintCheckCount, 0) & vbTab & _
        FormatCurrency(mcurCheckAmount, 2) & _
        vbCrLf & "Fees: " & vbTab & vbTab & vbTab & _
        FormatCurrency(mcurFeeAmount), _
        vbOKOnly, "Summary"
        
End Sub

Private Sub Form_Load()
    Show                    ' make form visible
    txtAmount.SetFocus      ' point to amount at start
End Sub
