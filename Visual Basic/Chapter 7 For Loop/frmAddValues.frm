VERSION 5.00
Begin VB.Form frmAddValues 
   AutoRedraw      =   -1  'True
   Caption         =   "Get total of 5 data values"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Average"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' program to get and add 5 values
Option Explicit

Private Sub cmdAverage_Click()
    Dim dblSum As Double    ' sum of values
    Dim dblValue As Double  ' input value
    Dim intLoop As Integer  ' to track which value
    Const intMax As Integer = 3 ' limit for loop
    For intLoop = 1 To intMax   ' loop through input
        dblValue = InputBox("Enter Value", "Average Data")
        dblSum = dblSum + dblValue
    Next ' end of For intLoop
    Print "The average is "; dblSum / intMax
End Sub

Private Sub cmdGetData_Click()
    ' need to get and add 5 values
    Dim dblSum As Double    ' sum of values
    Dim dblValue As Double  ' input value
    Dim intLoop As Integer  ' to track which value
    
    For intLoop = 1 To 5 Step 1    ' loop 5 times
        ' get a value
        dblValue = InputBox("Enter a Value", "Data Input")
        dblSum = dblSum + dblValue  ' add to sum
    Next  ' end of for loop
    Print "The sum is "; dblSum
End Sub

Private Sub cmdRandom_Click()
    Dim intSum As Integer    ' sum of values
    Dim intValue As Integer  ' input value
    Dim intLoop As Integer  ' to track which value
    
    For intLoop = 1 To 5    ' loop 5 times
        ' get a value
        intValue = Rnd() * 100
        intSum = intSum + intValue  ' add to sum
    Next  ' end of for loop
    Print "The sum of 5 random numbers is "; intSum

End Sub
