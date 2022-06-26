VERSION 5.00
Begin VB.Form prob31 
   AutoRedraw      =   -1  'True
   Caption         =   "Calorie Counter"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame fraTotals 
      Caption         =   "Totals"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   5055
      Begin VB.Label lblAll 
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblCurrent 
         Caption         =   "0"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Total Calories:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Items:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Current Calories:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraEntry 
      Caption         =   "Entry"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtProtein 
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCarbo 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFat 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "grams Protein"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "grams Carbohydrates"
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "grams Fat"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add current to total"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "prob31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 3.1 page 120
' Douglas Nielson
Option Explicit
Dim mintCount As Integer
Private mdblTotal As Double


Private Sub cmdAdd_Click()
    mdblTotal = mdblTotal + Val(lblCurrent.Caption)
    mintCount = mintCount + 1   ' add one to count
    lblCount.Caption = mintCount
    lblAll = mdblTotal
    ' clear entry
    txtFat.Text = "0"
    txtCarbo.Text = "0"
    txtProtein.Text = "0"
    lblCurrent.Caption = 0
    txtFat.SetFocus         ' moves cursor to box
End Sub

Private Sub cmdClear_Click()
    mintCount = 0           ' clear count
    mdblTotal = 0#          ' clear total
    lblCount = "0"
    lblAll = "0"
    txtFat.SetFocus         ' moves cursor to box
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    ' X and Y are twips into the button from the upper left corner as 0,0
    'Print "X="; X; "   Y="; Y
    If X < cmdPrint.Width / 2 Then ' coming from left side
        cmdPrint.Left = cmdPrint.Left + X
    Else    ' from right side
        cmdPrint.Left = cmdPrint.Left - (cmdPrint.Width - X)
    End If
    If Y < cmdPrint.Height / 2 Then ' coming from top
        cmdPrint.Top = cmdPrint.Top + Y
    Else    ' from bottom
        cmdPrint.Top = cmdPrint.Top - (cmdPrint.Height - Y)
    End If
    ' now if the button has gone off a side, wrap
    Const intBuffer As Integer = 200      ' don't reposition right on edge
    ' check top and bottom
    If cmdPrint.Top <= 0 Then ' off the top
        ' 510 is height in twips for the title bar across top of window
        cmdPrint.Top = Me.Height - cmdPrint.Height - intBuffer - 510
    ElseIf cmdPrint.Top >= Me.Height - cmdPrint.Height - 510 Then ' off bottom
        cmdPrint.Top = intBuffer
    End If
    ' check for left and right
    If cmdPrint.Left <= 0 Then
        cmdPrint.Left = Me.Width - cmdPrint.Width - intBuffer
    ElseIf cmdPrint.Left >= Me.Width - cmdPrint.Width Then ' off right
        cmdPrint.Left = intBuffer
    End If
End Sub

Private Sub txtCarbo_Change()
    If Len(txtFat.Text) = 0 Or Len(txtCarbo.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblCurrent = 9 * txtFat.Text + 4 * txtCarbo.Text + 4 _
        * txtProtein.Text
End Sub

Private Sub txtFat_Change()
    If Len(txtFat.Text) = 0 Or Len(txtCarbo.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblCurrent = 9 * txtFat.Text + 4 * txtCarbo.Text + 4 _
        * txtProtein.Text
End Sub

Private Sub txtProtein_Change()
    If Len(txtFat.Text) = 0 Or Len(txtCarbo.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblCurrent = 9 * txtFat.Text + 4 * txtCarbo.Text + 4 _
        * txtProtein.Text
End Sub
