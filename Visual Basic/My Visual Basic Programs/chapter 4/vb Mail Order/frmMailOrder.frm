VERSION 5.00
Begin VB.Form frmMailOrder 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   5055
      Begin VB.Line Line5 
         X1              =   240
         X2              =   4920
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblTaxTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbltaxShipping 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   4920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   4920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lbltaxSales 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3960
         TabIndex        =   36
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblTaxDollar 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblNonTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblNonShipping 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   4200
         X2              =   4800
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   3360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblNonSales 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblNonDollarAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxable"
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Non Taxable"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Shipping and Handling:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Sales Tax:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dollar Amount Due:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdNextItem 
      Caption         =   "Next Item"
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdateSummary 
      Caption         =   "Update Summary"
      Height          =   495
      Left            =   1320
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraItemOrdered 
      Caption         =   "Item Ordered"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   3720
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtWeight 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Weight:"
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   405
      End
   End
   Begin VB.Frame fraCustomerInfo 
      Caption         =   "Customer Info"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   480
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Zip:"
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "City:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdblSumWeight As Double ' to hold weight of order
Private mcurSumPrice As Currency
Const mdblSalesTax As Double = 0.08
Const mcurShipping As Currency = 0.25 ' cost per pound
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdNextItem_Click()
    ' for next item
    ' verify that quantity, weight and price are entered and valid number
    'verify description is entered
    ' verify customer info has been entered
    ' accumulate order totals of amount due and weight
    If Len(txtName.Text) <= 0 Then
        MsgBox "No Name Entered", vbOKOnly, "Error"
        txtName.SetFocus
        Exit Sub
    End If
    If Len(txtAddress.Text) <= 0 Then
        MsgBox "No Address Entered", vbOKOnly, "Error"
        txtAddress.SetFocus
        Exit Sub
    End If
    If Len(txtCity.Text) <= 0 Then
        MsgBox "No City Entered", vbOKOnly, "Error"
        txtCity.SetFocus
        Exit Sub
    End If
    If Len(txtState.Text) <= 0 Then
        MsgBox "No State Entered", vbOKOnly, "Error"
        txtState.SetFocus
        Exit Sub
    End If
    If Len(txtZip.Text) <= 0 Then
        MsgBox "No Zip Code Entered", vbOKOnly, "Error"
        txtZip.SetFocus
        Exit Sub
    End If
    If Len(txtDescription.Text) <= 0 Then
        MsgBox "No Description Entered", vbOKOnly, "Error"
        txtDescription.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtQuantity.Text) Then
        MsgBox "Non Numeric value for Quantity", vbOKOnly, "error"
        txtQuantity.SetFocus
        End If
    If Not IsNumeric(txtWeight.Text) Then
        MsgBox "Non Numeric value for Weight", vbOKOnly, "error"
        txtWeight.SetFocus
        End If
    If Not IsNumeric(txtPrice.Text) Then
        MsgBox "Non Numeric value for Price", vbOKOnly, "error"
        txtPrice.SetFocus
        Exit Sub
    End If
    ' add weight of this item to sum
    mdblSumWeight = mdblSumWeight + Val(txtQuantity.Text) * Val(txtWeight.Text)
    mcurSumPrice = mcurSumPrice + Val(txtPrice.Text) * Val(txtQuantity.Text)
    ' clear fields for next item
    txtDescription.Text = ""
    txtQuantity.Text = "1"
    txtWeight.Text = ""
    txtPrice.Text = ""
    txtDescription.SetFocus
End Sub
Private Sub cmdPrint_Click()
    MsgBox "NO PRINTING!", vbOKOnly, "Print Failure"
End Sub
Private Sub cmdUpdateSummary_Click()
    Dim curSalesTax As Currency
    Dim CurShipping As Currency
    Dim CurHandling As Currency
    curSalesTax = mdblSalesTax * mcurSumPrice
    CurShipping = mcurShipping * mdblSumWeight
    If mdblSumWeight < 10# Then
        CurHandling = 1
    ElseIf mdblSumWeight <= 100# Then
        CurHandling = 3
    Else
        CurHandling = 5#
    End If
    lblNonDollarAmount.Caption = FormatCurrency(mcurSumPrice)
    lblTaxDollarAmount.Caption = FormatCurrency(mcurSumPrice)
    lbltaxSales.Caption = FormatCurrency(curSalesTax)
    lblNonShipping.Caption = FormatCurrency(CurShipping + CurHandling)
End Sub

Private Sub lblNonSales_Click()

End Sub
