VERSION 5.00
Begin VB.Form frmSwitcher 
   Caption         =   "The Switcher"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00FFFF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame fraTextColor 
      Caption         =   "TextColor"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      Begin VB.OptionButton optGreen 
         Caption         =   "Green"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optRed 
         Caption         =   "Red"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optBlack 
         Caption         =   "Black"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image imgOn 
      Height          =   480
      Left            =   600
      Picture         =   "frmSwitcher.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Click to Turn Off"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Programmer Douglas Nielson"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      Caption         =   "Turn the light on"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image imgOff 
      Height          =   1080
      Left            =   2160
      Picture         =   "frmSwitcher.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "Click to turn on"
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmSwitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Douglas Nielson
' Problem 2.1 on page 80
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    ' disable for now
    MsgBox "No Print for Now", vbOKOnly, "Print Option"
End Sub

Private Sub Form_Load()
    ' allign the two images
    imgOn.Top = imgOff.Top
    imgOn.Left = imgOff.Left
    imgOn.Width = imgOff.Width
    imgOn.Height = imgOff.Height
    imgOn.Visible = False
End Sub

Private Sub imgOff_Click()
    imgOn.Visible = True
    imgOff.Visible = False
    lblMessage.Caption = "Turn the light off " & txtName.Text
End Sub

Private Sub imgOn_Click()
    imgOn.Visible = False
    imgOff.Visible = True
    lblMessage.Caption = "Turn the light on " & txtName.Text
End Sub

Private Sub lblMessage_Click()
    If imgOff.Visible = True Then  ' bulb is off
        imgOff.Visible = False
        imgOn.Visible = True
        lblMessage.Caption = "Turn the light off " & txtName.Text
    Else                            ' bulb is on
        imgOff.Visible = True
        imgOn.Visible = False
        lblMessage.Caption = "Turn the light on " & txtName.Text
    End If
End Sub

Private Sub optBlack_Click()
    lblMessage.ForeColor = vbBlack
End Sub

Private Sub optBlue_Click()
    lblMessage.ForeColor = vbBlue
End Sub

Private Sub optGreen_Click()
    lblMessage.ForeColor = vbGreen
End Sub

Private Sub optRed_Click()
    lblMessage.ForeColor = vbRed
End Sub
