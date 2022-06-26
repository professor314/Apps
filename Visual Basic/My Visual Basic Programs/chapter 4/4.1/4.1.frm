VERSION 5.00
Begin VB.Form frm41 
   Caption         =   "Problem number 4.1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextPatron 
      Caption         =   "&Next Patron"
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "&Summary"
      Height          =   495
      Left            =   1320
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton optManicure 
      Caption         =   "Manicure $35"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton OptHairStyling 
      Caption         =   "Hair Styling $60"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
   Begin VB.OptionButton optMakeover 
      Caption         =   "Make Over $125"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   2175
   End
   Begin VB.OptionButton optPermanentMakeup 
      Caption         =   "Permanent Makeup $200"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Services"
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discount"
      Height          =   1935
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton opt20Percent 
         Caption         =   "20% Discount"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton opt10Percent 
         Caption         =   "10% Discount"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optNoDiscount 
         Caption         =   "No Discount"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "C&alculate"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblTotalFees 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Fees:"
      Height          =   195
      Left            =   225
      TabIndex        =   12
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label lblfee 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Service Fee:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   900
   End
End
Attribute VB_Name = "frm41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sean Connolly
' Problem 4.1
' 1-27-02
Option Explicit
Private mcurTotalFees As Currency

Private Sub cmdCalculate_Click()
    Dim curfee As Currency
    If optMakeover.Value = True Then
        If optNoDiscount.Value = True Then
            curfee = 125#
        ElseIf opt10Percent = True Then
            curfee = 0.9 * 125#
        Else
            curfee = 0.8 * 125#
        End If
    ElseIf OptHairStyling.Value = True Then
        If optNoDiscount.Value = True Then
            curfee = 60#
        ElseIf opt10Percent = True Then
            curfee = 0.9 * 60#
        Else
            curfee = 0.8 * 60#
        End If
    ElseIf optManicure.Value = True Then
        If optNoDiscount.Value = True Then
            curfee = 35#
        ElseIf opt10Percent = True Then
            curfee = 0.9 * 35#
        Else
            curfee = 0.8 * 35#
        End If
    ElseIf optPermanentMakeup.Value = True Then
        If optNoDiscount.Value = True Then
            curfee = 200#
        ElseIf opt10Percent = True Then
            curfee = 0.9 * 200#
        Else
            curfee = 0.8 * 200#
        End If
    
    End If
    lblfee.Caption = FormatCurrency(curfee)
    mcurTotalFees = mcurTotalFees + curfee
    lblTotalFees = FormatCurrency(mcurTotalFees)
End Sub

Private Sub cmdClear_Click()
    lblfee.Caption = ""
    lblTotalFees.Caption = FormatCurrency(mcurTotalFees)
    optNoDiscount.Value = True
    mcurTotalFees = 0#
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "no print for now", vbOKOnly, "print control"
End Sub


