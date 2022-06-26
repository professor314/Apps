VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   1770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkProgramer 
      Caption         =   "Show Programer"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "by Sean Connolly"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkProgramer_Click()
    If chkProgramer.Value = vbChecked Then
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If
End Sub
