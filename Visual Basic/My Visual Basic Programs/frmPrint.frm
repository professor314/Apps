VERSION 5.00
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Output (to form)"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSoSO 
      Caption         =   "So So"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Average"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreat 
      Caption         =   "GREAT"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Rate Mr. Nielson as an Instructor"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Response to Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAverage_Click()
    cmdGreat.SetFocus
    cmdAverage.Value = False
    MsgBox "Wrong Answer, try again", vbOKOnly, "Wrong"
End Sub


Private Sub cmdAverage_LostFocus()
    cmdAverage.Value = True
End Sub

Private Sub cmdGreat_Click()
    frmMain.SetFocus    ' focus to other form
    Unload frmPrint     ' remove from memory
End Sub

Private Sub cmdSoSO_Click()
    MsgBox "Your Final Grade: B-", vbOKOnly, "Grade Change"
End Sub
