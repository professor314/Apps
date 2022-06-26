VERSION 5.00
Begin VB.Form frmHello 
   Caption         =   "Hello World Program"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdItalian 
      Caption         =   "&Italian"
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdEnglish 
      Caption         =   "&English"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdGerman 
      Caption         =   "&German"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "by Sean Connolly"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hello World Program"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Select Language"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program 1.1 Page 47
' By Sean Connolly
Option Explicit


Private Sub cmdEnglish_Click()
    ' Show Hello orld message in label
    lblMessage.Caption = "Hello World"
End Sub

Private Sub cmdExit_Click()
    ' end the program
    End
End Sub

Private Sub cmdGerman_Click()
    lblMessage.Caption = "Hallo Weld"
End Sub

Private Sub cmdItalian_Click()
    lblMessage.Caption = "Ciao Mondo Franch"
End Sub
