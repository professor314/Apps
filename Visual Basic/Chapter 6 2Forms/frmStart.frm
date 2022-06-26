VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Dividing Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' program to split form into two parts
Option Explicit

Private Sub cmdSplit_Click()
    Dim intWidth As Integer
    intWidth = frmStart.Width ' get width of form
    
    If cmdSplit.Caption = "Split" Then
        frmSecond.Top = frmStart.Top
        frmSecond.Left = frmStart.Left + intWidth / 2
        frmStart.Width = intWidth / 2  ' half as wide
        frmSecond.Width = frmStart.Width
        ' update split button
        frmStart.cmdSplit.Caption = "Unsplit"
        ' load the second form
        frmSecond.Show vbModeless
    Else  ' unsplit
        frmStart.Width = 2 * frmStart.Width
        frmSecond.Hide  ' disappear
        frmStart.cmdSplit.Caption = "Split"
    End If
End Sub
