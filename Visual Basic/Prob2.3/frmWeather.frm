VERSION 5.00
Begin VB.Form frmWeather 
   Caption         =   "Weather Report"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fraWeather 
      Caption         =   "Choose"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      Begin VB.OptionButton optSunny 
         Caption         =   "S&unny"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Choose if Sunny"
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optSnowy 
         Caption         =   "&Snowy"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Choose if Snowy"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optRainy 
         Caption         =   "&Rainy"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Choose if Rainy"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optCloudy 
         Caption         =   "&Cloudy"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Choose if Cloudy"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image imgSunny 
      Height          =   735
      Left            =   3240
      Picture         =   "frmWeather.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image imgSnowy 
      Height          =   735
      Left            =   1800
      Picture         =   "frmWeather.frx":0442
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image imgRainy 
      Height          =   735
      Left            =   3240
      Picture         =   "frmWeather.frx":0884
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image imgCloudy 
      Height          =   735
      Left            =   1800
      Picture         =   "frmWeather.frx":0CC6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Enter your name here"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' program 2.3
' by Douglas Nielson
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No Print for Now", vbOKOnly, "Print Request"
End Sub

Private Sub Form_Load()
    imgCloudy.Visible = False
    imgRainy.Visible = False
    imgSnowy.Visible = False
    imgSunny.Visible = False
End Sub

Private Sub optCloudy_Click()
    imgCloudy.Visible = True
    imgRainy.Visible = False
    imgSnowy.Visible = False
    imgSunny.Visible = False
    lblMessage.Caption = "It looks like cloudy weather today " _
        & txtName.Text
End Sub

Private Sub optRainy_Click()
    imgCloudy.Visible = False
    imgRainy.Visible = True
    imgSnowy.Visible = False
    imgSunny.Visible = False
    lblMessage.Caption = "It looks like rainy weather today " _
        & txtName.Text
End Sub

Private Sub optSnowy_Click()
    imgCloudy.Visible = False
    imgRainy.Visible = False
    imgSnowy.Visible = True
    imgSunny.Visible = False
    lblMessage.Caption = "It looks like snowy weather today " _
        & txtName.Text
End Sub

Private Sub optSunny_Click()
    imgCloudy.Visible = False
    imgRainy.Visible = False
    imgSnowy.Visible = False
    imgSunny.Visible = True
    lblMessage.Caption = "It looks like sunny weather today " _
        & txtName.Text
End Sub
