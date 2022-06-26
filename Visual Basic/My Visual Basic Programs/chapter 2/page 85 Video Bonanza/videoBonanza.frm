VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   FillColor       =   &H80000013&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Category"
      Height          =   1455
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton optAction 
         Caption         =   "&Action"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optHorror 
         Caption         =   "&Horror"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optScifi 
         Caption         =   "&Sci-Fi"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optDrama 
         Caption         =   "&Drama"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optComedy 
         Caption         =   "C&omedy"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lblAisle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aisle Number"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkmessage 
      Caption         =   "Show Members' Secret Message"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblTitle 
      Caption         =   "Video Bonanza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sean Connolly
' Page 85 Video Bonanza
' 1-20-03
Option Explicit

Private Sub chkmessage_Click()
    If chkmessage.Value = 1 Then
    lblMessage.Caption = "All Members Receive a 10% Discount!"
    Else
    lblMessage.Caption = ""
    End If
End Sub

Private Sub cmdClear_Click()
    lblMessage.Caption = ""
    optComedy.Value = False
    optDrama.Value = False
    optAction.Value = False
    optScifi.Value = False
    optHorror.Value = False
    chkmessage.Value = 0
    cmdClear.SetFocus
    lblAisle.Caption = "Aisle Number"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "no print for now", vbOKOnly, "print Request"
End Sub

Private Sub optAction_Click()
    lblAisle.Caption = "Aisle 3"
End Sub

Private Sub optComedy_Click()
    lblAisle.Caption = "Aisle 1"
End Sub

Private Sub optDrama_Click()
    lblAisle.Caption = "Aisle 2"
End Sub

Private Sub optHorror_Click()
    lblAisle.Caption = "Aisle 5"
End Sub

Private Sub optScifi_Click()
    lblAisle.Caption = "Aisle 4"
End Sub
