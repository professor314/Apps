VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sean's Info Program"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdCompany 
      Caption         =   "&Company"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdSchool 
      Caption         =   "&School"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "&Message"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Caption         =   "by Sean Connolly"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Sean's Info Program"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "Choose An Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program 1.2
' by Sean Connolly
Option Explicit

Private Sub cmdCompany_Click()
    ' displays company name
    lblMessage.Caption = "Skillings-Connolly, Inc."
End Sub

Private Sub cmdExit_Click()
    ' cancels program
    End
End Sub

Private Sub cmdMessage_Click()
    'displays greeting
    lblMessage.Caption = "Help! I'm trapped in the computer!"
End Sub

Private Sub cmdSchool_Click()
    ' displays college name
    lblMessage.Caption = "SPSCC"
End Sub

