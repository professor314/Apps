VERSION 5.00
Begin VB.Form frmproject31 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraentry 
      Caption         =   "Entry"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5175
      Begin VB.TextBox txtfat 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtcarbos 
         Height          =   285
         Left            =   1980
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtProtein 
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Grams Protein"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Grams Carbos"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Grams Fat"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fratotals 
      Caption         =   "Totals"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5175
      Begin VB.Label lblCalories 
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblcount 
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblcurrent 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Calories"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Number of Items"
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Current Calories"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdcurrent 
      Caption         =   "&Add Current to Total"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmproject31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 3.1
' 1-22-03
'Sean Connolly
Option Explicit
Dim mintcount As Integer
Private mdbltotal As Double

Private Sub cmdClear_Click()
    txtcarbos.Text = ""
    txtProtein.Text = ""
    txtfat.Text = ""
    lblcurrent.Caption = "0"
    lblcount.Caption = "0"
    lblCalories.Caption = "0"
    txtfat.SetFocus
End Sub

Private Sub cmdcurrent_Click()
    mdbltotal = mdbltotal + Val(lblcurrent.Caption)
    mintcount = mintcount + 1
    lblcount.Caption = mintcount
    lblCalories = mdbltotal
    ' clear entry
    txtfat.Text = ""
    txtcarbos.Text = ""
    txtProtein.Text = ""
    lblcurrent.Caption = 0
    txtfat.SetFocus        'moves cursor to box
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub txtcarbos_Change()
    If Len(txtfat.Text) = 0 Or Len(txtcarbos.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblcurrent.Caption = 9 * txtfat.Text + _
    4 * txtcarbos.Text + 4 * txtProtein.Text
End Sub

Private Sub txtfat_Change()
    If Len(txtfat.Text) = 0 Or Len(txtcarbos.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblcurrent.Caption = 9 * txtfat.Text + _
    4 * txtcarbos.Text + 4 * txtProtein.Text
End Sub

Private Sub txtProtein_Change()
    If Len(txtfat.Text) = 0 Or Len(txtcarbos.Text) = 0 Or _
        Len(txtProtein.Text) = 0 Then
        Exit Sub
    End If
    lblcurrent.Caption = 9 * txtfat.Text + _
    4 * txtcarbos.Text + 4 * txtProtein.Text
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
