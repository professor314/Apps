VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Summary"
      Height          =   1455
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Width           =   2535
      Begin VB.Label lblCustomers 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Customers Served:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Rental Income:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   720
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2535
      Begin VB.Label lblAmountDue 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblDiscount 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblRentalAmount 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Amount Due:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Discount:"
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rental Amount:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtMemberNumber 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Movies Rented:"
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Member Number:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub lblCustomers_Click()

End Sub
