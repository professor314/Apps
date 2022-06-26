VERSION 5.00
Begin VB.Form frmImage 
   Caption         =   "Image Consulting Shop Services"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDiscounts 
      Caption         =   "Discounts"
      Height          =   1455
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optTwenty 
         Caption         =   "20% discount"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optTen 
         Caption         =   "10% discount"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optZero 
         Caption         =   "no discount"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraServices 
      Caption         =   "Services"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optPermanent 
         Caption         =   "Permanent Makeup"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton optManicure 
         Caption         =   "Manicure"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optHairStyling 
         Caption         =   "Hair Styling"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optMakeover 
         Caption         =   "Makeover"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblTotalFees 
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Total Fees:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblFee 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Service Fee:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Douglas Nielson
' Problem 4.1 page 167
Option Explicit
Private mcurTotalFees As Currency

Private Sub cmdCalculate_Click()
    Dim curFee As Currency
    If optMakeover.Value = True Then 'Makeover is selected
        curFee = 125#
    ElseIf optHairStyling.Value = True Then ' Hair Styling
        curFee = 60#
    ElseIf optManicure.Value = True Then ' Manicure
        curFee = 35#
    ElseIf optPermanent.Value = True Then ' Permanent
        curFee = 200#
    End If ' end of Services option
    ' apply discount
    If optTen.Value = True Then
        curFee = 0.9 * curFee
    ElseIf optTwenty.Value = True Then
        curFee = 0.8 * curFee
    End If
    ' format output to form
    lblFee.Caption = FormatCurrency(curFee)
    mcurTotalFees = mcurTotalFees + curFee
    lblTotalFees = FormatCurrency(mcurTotalFees)
    optZero.Value = True
End Sub

Private Sub cmdClear_Click()
    lblFee.Caption = ""
    lblTotalFees.Caption = ""
    mcurTotalFees = 0#
    optZero.Value = True
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No print for now", vbOKOnly, "Print Control"
End Sub
