VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Picture"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image imgCD 
      Height          =   1455
      Left            =   1920
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgFloppy 
      Height          =   1320
      Left            =   480
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' douglas nielson
Option Explicit

Private Sub cmdChange_Click()
    imgFloppy.Visible = False
    imgCD.Visible = True
    
End Sub
