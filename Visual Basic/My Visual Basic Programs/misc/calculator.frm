VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Sean's Amazing Calculator"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdequals 
      Caption         =   "="
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmddivide 
      Caption         =   "/"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdtimes 
      Caption         =   "X"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H80000005&
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd0_Click()
    lblMessage.Caption = "0"
End Sub

Private Sub cmd1_Click()
    lblMessage.Caption = "1"
End Sub

Private Sub cmd2_Click()
    lblMessage.Caption = "2"
End Sub

Private Sub cmd3_Click()
    lblMessage.Caption = "3"
End Sub

Private Sub cmd4_Click()
    lblMessage.Caption = "4"
End Sub

Private Sub cmd5_Click()
    lblMessage.Caption = "5"
End Sub

Private Sub cmd6_Click()
    lblMessage.Caption = "6"
End Sub

