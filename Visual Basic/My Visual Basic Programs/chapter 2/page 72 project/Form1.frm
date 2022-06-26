VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1538
      TabIndex        =   15
      Top             =   3608
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Display"
      Height          =   375
      Left            =   578
      TabIndex        =   14
      Top             =   3608
      Width           =   855
   End
   Begin VB.Frame fraStyle 
      Caption         =   "Style"
      Height          =   1575
      Left            =   3120
      TabIndex        =   6
      Top             =   1928
      Width           =   1215
      Begin VB.CheckBox Check3 
         Caption         =   "&Underline"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "&Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "&Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   "Color"
      Height          =   1575
      Left            =   578
      TabIndex        =   5
      Top             =   1928
      Width           =   1215
      Begin VB.OptionButton optBlue 
         Caption         =   "&Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optYellow 
         Caption         =   "&Yellow"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optRed 
         Caption         =   "&Red"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   578
      TabIndex        =   0
      Top             =   735
      Width           =   3735
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Message:"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Label lblTitle 
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1560
      Left            =   1785
      TabIndex        =   13
      Top             =   1935
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdPrint_Click()
    MsgBox "no print for now", vbOKOnly, "print button"
End Sub

Private Sub lblTitle_Click()

End Sub
