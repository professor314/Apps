VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcost 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtend 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtstart 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblTurnovers 
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Turnovers"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Label lblAverage 
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Average Inventory"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost of Goods Sold"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ending Inventory"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Beginning Inventory"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 3.3
' 1-23-03
' Sean Connolly
Option Explicit
Private mcurAverageInventory As Currency
Dim mdblTurnovers As Double

Private Sub cmdClear_Click()
    txtstart = ""
    txtend = ""
    txtcost = ""
    lblAverage = ""
    lblTurnovers = ""
    txtstart.SetFocus
End Sub

Private Sub cmdCompute_Click()
    mcurAverageInventory = _
        (Val(txtstart.Text) + Val(txtend.Text)) / 2
    mdblTurnovers = Val(txtcost.Text) / mcurAverageInventory
    lblAverage = FormatCurrency(mcurAverageInventory, 2)
    lblTurnovers = FormatNumber(mdblTurnovers, 1)
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    cmdPrint.Visible = False
End Sub

