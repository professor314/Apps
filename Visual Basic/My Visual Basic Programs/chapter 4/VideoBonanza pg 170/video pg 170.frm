VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "New Release"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkMember 
      Caption         =   "Member"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movie Type"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1095
      Begin VB.OptionButton optVideo 
         Caption         =   "Video"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optDVD 
         Caption         =   "DVD"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "Howard the Duck"
      Top             =   555
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name of Movie:"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintNumberOfCutomers As Integer
