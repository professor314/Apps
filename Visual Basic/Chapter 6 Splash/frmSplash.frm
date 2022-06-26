VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H0000FFFF&
   Caption         =   "Welcome to My Program"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSplash 
      Interval        =   5000
      Left            =   2760
      Top             =   1920
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   240
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to CIS 160 Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' splash screen for example
Option Explicit

Private Sub tmrSplash_Timer()
    tmrSplash.Enabled = False   ' turn off timer
    ' Me refers to the current/active form
    Unload Me                   ' unload splash screen
    frmWork.Show vbModal        ' start other form
End Sub
