VERSION 5.00
Begin VB.Form frmFlagProgram 
   Caption         =   "Amazing Flag Program"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDisplay 
      Caption         =   "Display"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkformtitle 
         Caption         =   "Form Title"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkProgrammer 
         Caption         =   "Programmer"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkcountryname 
         Caption         =   "Country Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdKorea 
      Caption         =   "Korea"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanada 
      Caption         =   "Canada"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNorway 
      Caption         =   "Norway"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdusa 
      Caption         =   "USA"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpain 
      Caption         =   "Spain"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdTurkey 
      Caption         =   "Turkey"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdJapan 
      Caption         =   "Japan"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Flag Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   15
      Top             =   0
      Width           =   1425
   End
   Begin VB.Label lblProgrammer 
      AutoSize        =   -1  'True
      Caption         =   "by Sean Connolly"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1230
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image imgcanada 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgspain 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgjapan 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":0884
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgkorea 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":0CC6
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgnorway 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":1108
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgturkey 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":154A
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
   Begin VB.Image imgusa 
      Height          =   1215
      Left            =   2280
      Picture         =   "Form1.frx":198C
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1695
   End
End
Attribute VB_Name = "frmFlagProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project 2.2
'"Flags"'
' by Sean Connolly
Option Explicit

Private Sub chkcountryname_Click()
    lblCountry.Visible = False
End Sub

Private Sub chkformtitle_Click()
    lblTitle.Visible = False
End Sub

Private Sub chkProgrammer_Click()
    lblProgrammer.Visible = False
End Sub

Private Sub cmdCanada_Click()
    imgcanada.Visible = True
    imgusa.Visible = False
    imgjapan.Visible = False
    imgkorea.Visible = False
    imgnorway.Visible = False
    imgspain.Visible = False
    imgturkey.Visible = False
    lblCountry.Caption = "Canada"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdJapan_Click()
    imgcanada.Visible = False
    imgusa.Visible = False
    imgjapan.Visible = True
    imgkorea.Visible = False
    imgnorway.Visible = False
    imgspain.Visible = False
    imgturkey.Visible = False
    lblCountry.Caption = "Japan"
End Sub

Private Sub cmdKorea_Click()
    imgcanada.Visible = False
    imgusa.Visible = False
    imgjapan.Visible = False
    imgkorea.Visible = True
    imgnorway.Visible = False
    imgspain.Visible = False
    imgturkey.Visible = False
    lblCountry.Caption = "Korea"
End Sub

Private Sub cmdNorway_Click()
    imgcanada.Visible = False
    imgusa.Visible = False
    imgjapan.Visible = False
    imgkorea.Visible = False
    imgnorway.Visible = True
    imgspain.Visible = False
    imgturkey.Visible = False
    lblCountry.Caption = "Norway"
End Sub

Private Sub cmdPrint_Click()
    MsgBox "no print for now", vbOKOnly, "print Request"
End Sub

Private Sub cmdSpain_Click()
    imgcanada.Visible = False
    imgusa.Visible = False
    imgjapan.Visible = False
    imgkorea.Visible = False
    imgnorway.Visible = False
    imgspain.Visible = True
    imgturkey.Visible = False
    lblCountry.Caption = "Spain"
End Sub

Private Sub cmdTurkey_Click()
    imgcanada.Visible = False
    imgusa.Visible = False
    imgjapan.Visible = False
    imgkorea.Visible = False
    imgnorway.Visible = False
    imgspain.Visible = False
    imgturkey.Visible = True
    lblCountry.Caption = "Turkey"
End Sub

Private Sub cmdusa_Click()
    imgcanada.Visible = False
    imgusa.Visible = True
    imgjapan.Visible = False
    imgkorea.Visible = False
    imgnorway.Visible = False
    imgspain.Visible = False
    imgturkey.Visible = False
    lblCountry.Caption = "USA"
End Sub
