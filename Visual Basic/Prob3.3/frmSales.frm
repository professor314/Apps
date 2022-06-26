VERSION 5.00
Begin VB.Form frmSales 
   Caption         =   "Compute Inventory and Turnover"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtCost 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtEnd 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTurnovers 
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblAverage 
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Turnovers: "
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Average Inventory: "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cost of Goods Sold: "
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ending Inventory: "
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Beginning Inventory: "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program 3.3 on page 121
' Douglas Nielson
Option Explicit
Private mcurAverageInventory As Currency
Dim mdblTurnovers As Double

Private Sub cmdClear_Click()
    txtStart = ""
    txtEnd = ""
    txtCost = ""
    lblAverage = ""
    lblTurnovers = ""
    txtStart.SetFocus
End Sub

Private Sub cmdCompute_Click()
    mcurAverageInventory = _
        (Val(txtStart.Text) + Val(txtEnd.Text)) / 2
    mdblTurnovers = Val(txtCost.Text) / mcurAverageInventory
    lblAverage = FormatCurrency(mcurAverageInventory, 2)
    lblTurnovers = FormatNumber(mdblTurnovers, 1)
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No Print for now", vbOKOnly, "Print Control"
End Sub
