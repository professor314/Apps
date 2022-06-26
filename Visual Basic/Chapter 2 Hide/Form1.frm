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
   Begin VB.CheckBox chkProgrammer 
      Caption         =   "Show Programmer"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Label lblProgrammer 
      Caption         =   "programmed by Dilbert"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' sample to hide/restore an object
' Douglas Nielson
Option Explicit

Private Sub chkProgrammer_Click()
    If chkProgrammer.Value = vbChecked Then
        lblProgrammer.Visible = True
    Else
        lblProgrammer.Visible = False
    End If
End Sub
