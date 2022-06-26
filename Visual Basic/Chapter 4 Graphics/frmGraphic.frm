VERSION 5.00
Begin VB.Form frmGraphic 
   Caption         =   "Graphics"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Height          =   495
      Left            =   2040
      Picture         =   "frmGraphic.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Height          =   495
      Left            =   3480
      Picture         =   "frmGraphic.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Program"
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frmGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program showing graphics examples
' Douglas Nielson Jan. 31, 2003
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No Print for now", vbOKOnly, "Print Message"
End Sub
