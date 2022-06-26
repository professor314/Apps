VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPieceCalculator 
   Caption         =   "Piece Pay Calculator"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAmount 
      Caption         =   "$0.00"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblPay 
      Caption         =   "Pay Amount:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number of Pieces:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCalcpay 
         Caption         =   "&Calc Pay"
      End
      Begin VB.Menu mnuFileSummary 
         Caption         =   "&Summary"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuEditColor 
         Caption         =   "&Color"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmPieceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 5.1 page 207
' Douglas Nielson   Feb. 11, 2003
Option Explicit
Dim mintNumberOfPeople As Integer
Dim mlngNumberOfPieces As Long
Dim mcurTotalAmount As Currency

Function PayAmount(ByVal intNumPieces As Integer) As Currency
    ' compute the earnings for a person see table page 208
    Dim curReturn As Currency   ' value to be returned
    If intNumPieces < 1 Then
        curReturn = 0
    ElseIf intNumPieces <= 199 Then
        curReturn = 0.5 * intNumPieces
    ElseIf intNumPieces <= 399 Then
        curReturn = 0.55 * intNumPieces
    ElseIf intNumPieces <= 599 Then
        curReturn = 0.6 * intNumPieces
    Else
        curReturn = 0.65 * intNumPieces
    End If
    PayAmount = curReturn
End Function
Private Sub mnuEditClear_Click()
    ' clear input value
    txtNumber.Text = ""
    txtNumber.SetFocus
End Sub

Private Sub mnuEditColor_Click()
    ' change color of labels and text box
    ' from page 203 in text book
    With dlgCommon
        .Flags = cdlCCRGBInit   ' set initial color
        .Color = txtNumber.ForeColor ' to current color
        .ShowColor              ' dialog for color selection
        lblNumber.ForeColor = .Color
        txtNumber.ForeColor = .Color
        lblPay.ForeColor = .Color
        lblAmount.ForeColor = .Color
    End With
End Sub

Private Sub mnuEditFont_Click()
    ' change the font - see page 203 text book
    With dlgCommon
        .Flags = cdlCFScreenFonts ' Set Font dialog box fonts
        .ShowFont       ' runs the dialog
        lblNumber.Font.Name = .FontName
        lblNumber.Font.Size = .FontSize
                
        lblNumber.Font.Bold = .FontBold   ' assign selected
        lblNumber.Font.Italic = .FontItalic

        txtNumber.Font.Name = .FontName
        txtNumber.Font.Size = .FontSize
        lblPay.Font.Name = .FontName
        lblPay.Font.Size = .FontSize
        lblAmount.Font.Name = .FontName
        lblAmount.Font.Size = .FontSize
    End With
End Sub

Private Sub mnuFileCalcpay_Click()
    Dim curPay As Currency
    If IsNumeric(txtNumber.Text) Then
        curPay = PayAmount(Val(txtNumber.Text))
        lblAmount = FormatCurrency(curPay)
        ' update totals
        mintNumberOfPeople = mintNumberOfPeople + 1 ' add in one more person
        ' add pieces by this person to total number of pieces
        mlngNumberOfPieces = mlngNumberOfPieces + Val(txtNumber.Text)
        ' add pay to total pay
        mcurTotalAmount = mcurTotalAmount + curPay
    Else
        MsgBox "Number of Pieces is not Numeric", vbCritical + vbOKOnly, _
            "Error Message"
        txtNumber.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    ' end the program
    End
End Sub

Private Sub mnuFileSummary_Click()
    Dim strMessage As String        ' for message box string
    Dim curAverage As Currency
    If mintNumberOfPeople = 0 Then
        MsgBox "No data Entered", vbOKOnly, "Message"
        Exit Sub
    End If
    strMessage = "Total Number of Pieces is " & _
        Str(mlngNumberOfPieces) & vbCrLf
    strMessage = strMessage & "Total Pay is " & _
        FormatCurrency(mcurTotalAmount) & vbCrLf
    curAverage = mcurTotalAmount / mintNumberOfPeople
    strMessage = strMessage & "Average Pay is " & _
        FormatCurrency(curAverage)
    MsgBox strMessage, vbOKOnly + vbInformation, "Summary"
End Sub
