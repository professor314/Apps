VERSION 5.00
Begin VB.Form frmweatherreport 
   Caption         =   "Weather Report2.3"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame frachoose 
      Caption         =   "Choose"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
      Begin VB.OptionButton optSunny 
         Caption         =   "S&unny"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optSnowy 
         Caption         =   "&Snowy"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optRainy 
         Caption         =   "&Rainy"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optCloudy 
         Caption         =   "&Cloudy"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Your Name"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblMessage 
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Image imgsun 
      Height          =   1800
      Left            =   1440
      Picture         =   "frmweatherreport.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgsnow 
      Height          =   1800
      Left            =   1560
      Picture         =   "frmweatherreport.frx":0442
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgrain 
      Height          =   1800
      Left            =   1560
      Picture         =   "frmweatherreport.frx":0884
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgcloud 
      Height          =   1800
      Left            =   1560
      Picture         =   "frmweatherreport.frx":0CC6
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      Caption         =   "by Sean Connolly"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label lblEntername 
      AutoSize        =   -1  'True
      Caption         =   "Enter Your Name here:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmweatherreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Weather Report
' By Sean Connolly
Option Explicit
Private Sub cmdexit_Click()
    End
End Sub
Private Sub cmdprint_Click()
    MsgBox "no print for now", vbOKOnly, "print Request"
End Sub

Private Sub optCloudy_Click()
    imgcloud.Visible = True
    imgrain.Visible = False
    imgsun.Visible = False
    imgsnow.Visible = False
    lblMessage.Caption = "It looks like it will be cloudy today, " & txtName
End Sub

Private Sub optRainy_Click()
    imgcloud.Visible = False
    imgrain.Visible = True
    imgsun.Visible = False
    imgsnow.Visible = False
    lblMessage.Caption = "It looks like it will be rainy today, " & txtName
End Sub

Private Sub optSnowy_Click()
imgcloud.Visible = False
    imgrain.Visible = False
    imgsun.Visible = False
    imgsnow.Visible = True
    lblMessage.Caption = "It looks like it will be Snowy today, " & txtName
End Sub

Private Sub optSunny_Click()
    imgcloud.Visible = False
    imgrain.Visible = False
    imgsun.Visible = True
    imgsnow.Visible = False
    lblMessage.Caption = "It looks like it will be Sunny today, " & txtName
End Sub
