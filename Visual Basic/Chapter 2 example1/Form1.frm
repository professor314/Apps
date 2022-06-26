VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "&Order"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame fraFlavor 
      BackColor       =   &H0080FFFF&
      Caption         =   "Flavor Selection"
      Height          =   1575
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkStrawberry 
         BackColor       =   &H0080FFFF&
         Caption         =   "Strawberry"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkChocolate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Chocolate"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkVanilla 
         BackColor       =   &H0080FFFF&
         Caption         =   "Vanilla"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame fraScoops 
      Caption         =   "How Many Scoops"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optThree 
         Caption         =   "Three Scoops"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optTwo 
         Caption         =   "Two Scoops"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optOne 
         Caption         =   "One Scoop"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Douglas Nielson
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOrder_Click()
    Dim strOrder As String  ' define the variable
    If optOne.Value = True Then
        strOrder = "One Scoop"
    ElseIf optTwo.Value = True Then
        strOrder = "Two Scoops"
    Else
        strOrder = "Three Scoops"
    End If
    If chkVanilla.Value = vbChecked Then
        strOrder = strOrder & " Vanilla"
    End If
    If chkChocolate.Value = vbChecked Then
        strOrder = strOrder & " Chocolate"
    End If
    If chkStrawberry.Value = vbChecked Then
        strOrder = strOrder & " Strawberry"
    End If
    MsgBox strOrder, vbOKOnly, "Order"
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No print for now", vbOKOnly, "Print Button"
End Sub
