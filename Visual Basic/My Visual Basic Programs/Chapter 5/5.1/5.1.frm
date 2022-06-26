VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPieces 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   450
      Width           =   855
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblPay 
      AutoSize        =   -1  'True
      Caption         =   "Pay Amount:"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
   Begin VB.Label lblPieces 
      AutoSize        =   -1  'True
      Caption         =   "Number of Pieces:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCalcPay 
         Caption         =   "&Calc Pay"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileSummary 
         Caption         =   "&Summary"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFont 
         Caption         =   "&Font"
         Shortcut        =   ^F
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sean Connolly
' Problem 5.1
Option Explicit
Dim mintNumberOfPeople As Integer
Dim mlngNumberofPieces As Long
Dim mcurTotalAmount As Currency

Function PayAmount(ByVal intNumPieces As Integer) As Currency
    ' compute the earnings for a person see table page 208
    Dim curReturn As Currency  ' value to be returned
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
    txtPieces.Text = ""
    txtPieces.SetFocus
End Sub

Private Sub mnuEditColor_Click()
    ' Change colors of lables and text box
    With dlgCommon
        .Flags = cdlCCRGBInit 'set up initial color
        .Color = txtPieces.ForeColor
        .ShowColor ' Brings up dialog box
        lblAmount.ForeColor = .Color
        txtPieces.ForeColor = .Color
        lblPay.ForeColor = .Color
        lblPieces.ForeColor = .Color
    End With
End Sub

Private Sub mnuEditFont_Click()
    With dlgCommon
        .Flags = cdlCFScreenFonts
        .ShowFont
        lblAmount.Font.Name = .FontName
        txtPieces.Font.Name = .FontName
        lblPay.Font.Name = .FontName
        lblPieces.Font.Name = .FontName
    End With
End Sub

Private Sub mnuFileCalcPay_Click()
    Dim curPay As Currency
    If IsNumeric(txtPieces.Text) Then
        curPay = PayAmount(Val(txtPieces.Text))
        lblAmount = FormatCurrency(curPay)
        ' update totals
        mintNumberOfPeople = mintNumberOfPeople + 1
        mlngNumberofPieces = mlngNumberofPieces + Val(txtPieces.Text)
        mcurTotalAmount = mcurTotalAmount + curPay
    Else
        MsgBox "Number of Pieces is not Numberic", vbCritical + vbOKOnly, "error message"
        txtPieces.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileSummary_Click()
    Dim strMessage As String
    Dim curAverage As Currency
    If mintNumberOfPeople = 0 Then
        MsgBox "No Data Entered", vbOKOnly, "Message"
        Exit Sub
    End If
    strMessage = "Total Number of Pieces is " & Str(mlngNumberofPieces) & vbCrLf
    strMessage = strMessage & "Total Pay is " & FormatCurrency(mcurTotalAmount) & vbCrLf
    curAverage = mcurTotalAmount / mintNumberOfPeople
    strMessage = strMessage & "Average Pay is " & _
        FormatCurrency(curAverage)
    MsgBox strMessage, vbOKOnly + vbInformation, "Summary"
End Sub
