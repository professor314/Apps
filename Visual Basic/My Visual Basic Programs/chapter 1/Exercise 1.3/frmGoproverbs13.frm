VERSION 5.00
Begin VB.Form frmGoproverbs13 
   Caption         =   """Go"" Proverbs Program"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000003&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdConnecting 
      Caption         =   "&Connecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdMoves 
      Caption         =   "&Moves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdLife 
      Caption         =   "&Life"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdShape 
      Caption         =   "&Shape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      Caption         =   "by Sean Connolly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   3720
      Width           =   1230
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Go Proverb Program"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      Caption         =   "Choose a Proverb"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmGoproverbs13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Go Proverbs Program
' Project Exercise 1.3
' by Sean Connolly
Option Explicit

Private Sub cmdConnecting_Click()
    ' shows proverb for connecting
    lblDisplay.Caption = "The knight's jump is a guaranteed connection"
End Sub

Private Sub cmdExit_Click()
    ' cancel program
    End
End Sub

Private Sub cmdLife_Click()
    ' shows proverb for life
    lblDisplay.Caption = "Two eyes lives and one eye dies"
End Sub

Private Sub cmdMoves_Click()
    ' shows proverb for choosing a move
    lblDisplay.Caption = "Your best move is your opponents best move"
End Sub

Private Sub cmdShape_Click()
    ' shows proverb for shape
    lblDisplay.Caption = "Pon-noki is worth 30 points"
End Sub

