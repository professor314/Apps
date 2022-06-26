VERSION 5.00
Begin VB.Form frmSwitcher 
   BackColor       =   &H0080C0FF&
   Caption         =   "Problem 2.1, The Switcher"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "click to exit"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   2520
      MaskColor       =   &H0080C0FF&
      TabIndex        =   7
      ToolTipText     =   "click to print form"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "Your Name"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame fraColor 
      BackColor       =   &H0080C0FF&
      Caption         =   "Text Color"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "choose a color"
      Top             =   480
      Width           =   1335
      Begin VB.OptionButton optGreen 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Green"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optRed 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Red"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optBlue 
         BackColor       =   &H0080C0FF&
         Caption         =   "B&lue"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optBlack 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Black"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image imgoff 
      Height          =   1560
      Left            =   1800
      Picture         =   "2.1.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1560
   End
   Begin VB.Image imgon 
      Height          =   1560
      Left            =   1800
      Picture         =   "2.1.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "Click to turn off"
      Top             =   600
      Width           =   1560
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H0080C0FF&
      Caption         =   "Turn the Light on"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "by Sean Connolly"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1230
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmSwitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Problem 2.1
'by Sean Connolly
Option Explicit
Private Sub cmdExit_Click()
    'cancels Program
    End
End Sub

Private Sub cmdPrint_Click()
    'gives message about the print function
    MsgBox "No Print for Now", vbOKOnly, "Print Option"
End Sub

Private Sub imgoff_Click()
    'turn image on
    imgon.Visible = True
    imgoff.Visible = False
    lblMessage.Caption = "turn the light off " & txtName.Text
End Sub

Private Sub imgon_Click()
    imgon.Visible = False
    imgoff.Visible = True
    lblMessage.Caption = "turn the light on " & txtName.Text
End Sub

Private Sub lblMessage_Click()
  If imgoff.Visible = True Then
    imgoff.Visible = False
    imgon.Visible = True
    lblMessage.Caption = "turn the light off " & txtName
  Else
    imgon.Visible = False
    imgoff.Visible = True
    lblMessage.Caption = "turn the light on " & txtName
    End If

End Sub

Private Sub optBlack_Click()
    lblMessage.ForeColor = vbBlack
    lblAuthor.ForeColor = vbBlack
    lblName.ForeColor = vbBlack
    txtName.ForeColor = vbBlack
End Sub

Private Sub optBlue_Click()
    lblMessage.ForeColor = vbBlue
    lblAuthor.ForeColor = vbBlue
    lblName.ForeColor = vbBlue
    txtName.ForeColor = vbBlue
End Sub

Private Sub optGreen_Click()
    lblMessage.ForeColor = vbGreen
    lblAuthor.ForeColor = vbGreen
    lblName.ForeColor = vbGreen
    txtName.ForeColor = vbGreen
End Sub

Private Sub optRed_Click()
    lblMessage.ForeColor = vbRed
    lblAuthor.ForeColor = vbRed
    lblName.ForeColor = vbRed
    txtName.ForeColor = vbRed
End Sub
