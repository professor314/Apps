VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Summary"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   3735
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Average Pay:"
         Height          =   195
         Left            =   870
         TabIndex        =   13
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Pay:"
         Height          =   195
         Left            =   1110
         TabIndex        =   12
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Number of Pieces:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdDontTouch 
      Caption         =   "&Don't Touch"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "&Summary"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPiecesCompleted 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pieces Completed:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintNumberOfPeople As Integer

Private Sub cmdClear_Click()
    txtName.Text = ""
    txtPiecesCompleted.Text = "0"
    mintNumberOfPeople = mintNumberOfPeople + 1
End Sub

Private Sub cmdDontTouch_Click()
    MsgBox "It specifically told you not to touch it!", vbOKOnly, "Now You've Done it!"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

