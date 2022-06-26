VERSION 5.00
Begin VB.Form frmList 
   Caption         =   "Example of List Box"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboData 
      Height          =   315
      ItemData        =   "frmList.frx":0000
      Left            =   120
      List            =   "frmList.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "None Selected"
      Top             =   720
      Width           =   3135
   End
   Begin VB.ListBox lstData 
      Height          =   1815
      ItemData        =   "frmList.frx":0026
      Left            =   120
      List            =   "frmList.frx":0030
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' example of load a list box
Option Explicit

Private Sub cmdAdd_Click()
    Dim strWork As String
    strWork = txtAdd.Text   ' get from text box
    Trim (strWork)          ' remove leading trailing spaces
    If Len(strWork) > 0 Then    ' not an empty string
        lstData.AddItem (strWork)   ' add to list box
    End If  ' string has a non zero length
    ' clear text box and set focus
    txtAdd.Text = ""
    txtAdd.SetFocus
End Sub

Private Sub Form_Load()
    cboData.ListIndex = 0
    
End Sub

Private Sub lstData_Click()
    ' move selected item to text box
    txtAdd.Text = lstData.List(lstData.ListIndex)
End Sub
