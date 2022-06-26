VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAverage 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStdDev 
      Caption         =   "Standard Deviation"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Average"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lstData 
      Height          =   2595
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   1920
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "File Open"
      FileName        =   "MyData.Txt"
      Filter          =   "Text Files|*.txt|All Files|*.*"
   End
End
Attribute VB_Name = "frmAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' program to read data from a text file into a list box
'   find the average
'   find the standard deviation
Option Explicit
Function Average() As Double
    ' find the average of the data
    Dim intSum As Integer       ' sum of data items
    Dim intLoop As Integer      ' to loop through list box
    ' there are ListCount values in list box
    For intLoop = 0 To lstData.ListCount - 1
        intSum = intSum + Val(lstData.List(intLoop))
    Next
    ' average is sume / count
    Average = CDbl(intSum) / CDbl(lstData.ListCount)
End Function
Private Sub cmdAverage_Click()
    Print "The average is "; FormatNumber(Average(), 1)
End Sub

Private Sub cmdStdDev_Click()
    ' find standard deviation
    Dim intLoop As Integer
    Dim dblSum As Double
    Dim dblAverage As Double
    Dim dblStdDev As Double
    dblAverage = Average()      ' call average function
    ' need sum of ( x - average) ^ 2
    For intLoop = 0 To lstData.ListCount - 1
        dblSum = dblSum + (Val(lstData.List(intLoop)) - dblAverage) ^ 2
    Next
    dblStdDev = Sqr(dblSum / (CDbl(lstData.ListCount) - 1#))
    Print "The standard deviation is "; FormatNumber(dblStdDev, 1)
End Sub

Private Sub Form_Load()
    ' get data from a text file
    Dim strData As String   ' to hold data from file
    cdlOpen.ShowOpen        ' show open dial box
    Open cdlOpen.FileName For Input As #1  ' link to file
    Do While Not EOF(1)
        Input #1, strData   ' get some data from file
        Trim (strData)      ' get rid of spaces
        If Len(strData) > 0 Then ' not empty data
            lstData.AddItem (strData)   ' add to list box
        End If
    Loop
End Sub
