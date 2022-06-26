VERSION 5.00
Begin VB.Form frmStates 
   Caption         =   "States and Territories"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstState 
      Height          =   3180
      ItemData        =   "frmStates.frx":0000
      Left            =   240
      List            =   "frmStates.frx":00A9
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "List Box Index"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblIndex 
      Caption         =   "-1"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "State: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 7.4 page 288
Option Explicit

Private Sub lstState_Click()
    ' just a quick way to show the selected index
    lblIndex.Caption = Str(lstState.ListIndex)
End Sub

Private Sub txtState_Change()
    ' when the box changes then try to find a match in lstState
    ' lstState.ListCount has number of entries in the list box
    ' lstState.List( position ) has the actual string entry
    ' lstState.ListIndex is the position of the selected entry
    Dim intNumberOfCharacters As Integer    ' characters to compart from text box
    Dim strFromTextBox As String    ' to use for compare
    
    ' get string from Text box
    strFromTextBox = Trim(txtState.Text)    ' get copy
    intNumberOfCharacters = Len(strFromTextBox) ' how many characters
    strFromTextBox = UCase(strFromTextBox)  ' force all uppper case
    ' if no characters then unselect everyting
    If intNumberOfCharacters = 0 Then
        lstState.ListIndex = -1  ' unselect from list box
        Exit Sub    ' exit subroutine
    End If
    ' search to find item in list and select
    ' if not exact match, select Item before position
    ' use a loop to find match
    Dim intLoop As Integer          ' for search loop
    Dim strFromListBox As String    ' to make testing simplier
    For intLoop = 0 To lstState.ListCount - 1 ' zero based
        strFromListBox = UCase(lstState.List(intLoop))
        ' get same length as string in txtState box
        strFromListBox = Left(strFromListBox, intNumberOfCharacters)
        If strFromListBox = strFromTextBox Then ' have match
            lstState.ListIndex = intLoop    ' select item
            Exit For                ' done with loop
        ElseIf strFromListBox > strFromTextBox Then ' gone too far and no match
            lstState.ListIndex = intLoop - 1 ' previous item is best
            Exit For                ' done with loop
        ElseIf intLoop = lstState.ListCount - 1 Then  ' at last entry and not found
            lstState.ListIndex = intLoop
        End If
    Next    ' end of for loop
End Sub
