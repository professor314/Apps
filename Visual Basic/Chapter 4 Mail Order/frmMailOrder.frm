VERSION 5.00
Begin VB.Form frmMailOrder 
   Caption         =   "VB Mail Order"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3960
      TabIndex        =   38
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2760
      TabIndex        =   37
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Summary"
      Height          =   495
      Left            =   1320
      TabIndex        =   36
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextItem 
      Caption         =   "&Next Item"
      Height          =   495
      Left            =   0
      TabIndex        =   35
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   3120
      Width           =   4935
      Begin VB.Label lblTTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblNTTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Total:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblTShip 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblNTShip 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Shipping and Handling:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblTSalesTax 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblNTSalesTax 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Sales Tax:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblTDue 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblNTDue 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxable"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Non Taxable"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Dollar Amount Due:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraItemOrdered 
      Caption         =   "Item Ordered"
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   4935
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtWeight 
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtDescription 
         Height          =   405
         Left            =   1080
         TabIndex        =   16
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label9 
         Caption         =   "Price: "
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Weight:"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Description: "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraCustomerInfo 
      Caption         =   "Customer Info"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtZip 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtState 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Zip: "
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "State: "
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "City:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB Mail Order page 169
' Douglas Nielson    Feb 4-5, 2003
Option Explicit
Private mdblSumWeight As Double ' to hold weight of order
Private mcurSumPrice As Currency ' to hold price of order
Const mdblSaleTax As Double = 0.08 ' CA sales tax
Const mcurShipping As Currency = 0.25 ' cost per pound

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdNextItem_Click()
    ' for next item
    ' verify that Quantity, Weight, Price are
    '    entered and valid number
    ' verify Description is entered
    ' verify Customer Info has been entered
    ' accumulate order totals of
    '   amount Due, Weight
    If Len(txtName.Text) <= 0 Then ' no name entered
        MsgBox "No Name Entered", vbOKOnly, "Error"
        txtName.SetFocus
        Exit Sub
    End If
    If Len(txtAddress.Text) <= 0 Then ' no address entered
        MsgBox "No Address Entered", vbOKOnly, "Error"
        txtAddress.SetFocus
        Exit Sub
    End If
    If Len(txtCity.Text) <= 0 Then ' no city entered
        MsgBox "No City Entered", vbOKOnly, "Error"
        txtCity.SetFocus
        Exit Sub
    End If
    If Len(txtState.Text) <= 0 Then ' no state entered
        MsgBox "No State Entered", vbOKOnly, "Error"
        txtState.SetFocus
        Exit Sub
    End If
    If Len(txtZip.Text) <= 0 Then ' no zipcode entered
        MsgBox "No Zip Code Entered", vbOKOnly, "Error"
        txtZip.SetFocus
        Exit Sub
    End If
    If Len(txtDescription.Text) <= 0 Then ' no description entered
        MsgBox "No Description Entered", vbOKOnly, "Error"
        txtDescription.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtQuantity.Text) Then
        MsgBox "Non Numberic value for Quantity", _
            vbOKOnly, "Error"
        txtQuantity.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtWeight.Text) Then
        MsgBox "Non Numberic value for Weight", _
            vbOKOnly, "Error"
        txtWeight.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtPrice.Text) Then
        MsgBox "Non Numberic value for Price", _
            vbOKOnly, "Error"
        txtPrice.SetFocus
        Exit Sub
    End If
    ' add weight of this item to sum
    mdblSumWeight = mdblSumWeight + _
        Val(txtWeight.Text) * Val(txtQuantity.Text)
    mcurSumPrice = mcurSumPrice + _
        Val(txtPrice.Text) * Val(txtQuantity.Text)
    ' clear fields for next item
    txtDescription.Text = ""
    txtQuantity.Text = ""
    txtWeight.Text = ""
    txtPrice.Text = ""
    txtDescription.SetFocus
End Sub

Private Sub cmdPrint_Click()
    MsgBox "No Print for Now", vbOKOnly, "Print Message"
End Sub

Private Sub cmdUpdate_Click()
    Dim curSalesTax As Currency
    Dim curShipping As Currency
    Dim curHandling As Currency
    
    If txtDescription.Text <> "" Then
        Call cmdNextItem_Click
    End If
    curSalesTax = mdblSaleTax * mcurSumPrice
    curShipping = mcurShipping * mdblSumWeight
    If mdblSumWeight < 10# Then
        curHandling = 1#
    ElseIf mdblSumWeight <= 100# Then
        curHandling = 3#
    Else
        curHandling = 5#
    End If
    lblNTDue.Caption = FormatCurrency(mcurSumPrice)
    lblTDue.Caption = FormatCurrency(mcurSumPrice)
    lblTSalesTax.Caption = FormatCurrency(curSalesTax)
    lblNTShip.Caption = FormatCurrency(curShipping + _
        curHandling)
    lblTShip.Caption = FormatCurrency(curShipping + _
        curHandling)
    ' amount due
    lblNTTotal = FormatCurrency(mcurSumPrice + _
        curShipping + curHandling)
    lblTTotal = FormatCurrency(mcurSumPrice + _
        curShipping + curHandling + curSalesTax)
End Sub
