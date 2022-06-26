VERSION 5.00
Begin VB.Form frmproblem32 
   Caption         =   "Lennie's Bail Bonds"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCollateral 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtBail 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "C&alculate"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblFee 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fee"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   840
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Collateral"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bail Amount"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmproblem32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sean Connolly
' Problem 3.2
' 1-22-03
Option Explicit

Private Sub cmdCalculate_Click()
    Dim curfee As Currency
    Dim curBail As Currency
    curBail = Val(txtBail)
    curfee = Val(curBail) * (0.1)
    lblFee.Caption = FormatCurrency(curfee)
End Sub

Private Sub cmdClear_Click()
    lblFee.Caption = ""
    txtCollateral.Text = ""
    txtBail.Text = ""
    txtBail.SetFocus
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "no print for now", vbOKOnly, "Print Request"
End Sub
