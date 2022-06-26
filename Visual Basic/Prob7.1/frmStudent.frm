VERSION 5.00
Begin VB.Form frmStudent 
   Caption         =   "Student Information"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDeanList 
      Caption         =   "Dean's List"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox cboHighSchool 
      Height          =   315
      ItemData        =   "frmStudent.frx":0000
      Left            =   1320
      List            =   "frmStudent.frx":0010
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   4200
      Width           =   3015
   End
   Begin VB.ListBox lstMajor 
      Height          =   840
      ItemData        =   "frmStudent.frx":0040
      Left            =   1320
      List            =   "frmStudent.frx":0050
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Frame fraClass 
      Caption         =   "Class"
      Height          =   1575
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
      Begin VB.OptionButton optSenior 
         Caption         =   "Senior"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optJunior 
         Caption         =   "Junior"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optSophomore 
         Caption         =   "Sophomore"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optFreshman 
         Caption         =   "Freshman"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtUnits 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "High School: "
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Majors:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Units Complete: "
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrintSchool 
         Caption         =   "&Print School"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Problem 7.1 page 286
' By Douglas Nielson
Option Explicit
Public mstrTypePrint As String ' global to access in other form

Private Sub cmdOk_Click()
    ' clear fields and setup for next student
    Dim intLoop As Integer      ' loop variable
    Dim blnFound As Boolean     ' test for highschool in list
    
    blnFound = False ' found a High School match
    txtName.Text = ""
    txtUnits.Text = ""
    chkDeanList.Value = False
    optFreshman.Value = True
    lstMajor.ListIndex = -1  ' remove selection
    ' see if high school is already in list
    For intLoop = 0 To cboHighSchool.ListCount - 1
        If cboHighSchool.List(intLoop) = cboHighSchool.Text Then
            blnFound = True
            Exit For        ' stop loop
        End If
    Next
    If blnFound = False Then ' add to list
        cboHighSchool.AddItem (cboHighSchool.Text)
    End If
    cboHighSchool.Text = ""     ' clear entry
    
End Sub

Private Sub cmdPrint_Click()    ' print student info
    mstrTypePrint = "Student"
    frmPrint.Show
    Me.Hide
    frmPrint.Cls        ' clear old printing
    
    frmPrint.FontSize = 10
    frmPrint.FontName = "Arial"
    frmPrint.Print vbTab; "Student Summary"
    frmPrint.Print vbTab; "Name: "; txtName.Text
    frmPrint.Print vbTab; "Units Complete: "; txtUnits.Text
    If chkDeanList.Value = vbUnchecked Then
        frmPrint.Print vbTab; "NOT ";
    Else
        frmPrint.Print vbTab;
    End If
    frmPrint.Print "On Dean's List"
    frmPrint.Print vbTab; "Class: ";
    If optFreshman.Value = True Then
        frmPrint.Print "Freshman"
    ElseIf optSophomore.Value = True Then
        frmPrint.Print "Sophomore"
    ElseIf optJunior.Value = True Then
        frmPrint.Print "Junior"
    Else
        frmPrint.Print "Senior"
    End If
    frmPrint.Print vbTab; "Major: ";
    If lstMajor.ListIndex = -1 Then
        frmPrint.Print "None Selected"
    Else
        frmPrint.Print lstMajor.List(lstMajor.ListIndex)
    End If
    frmPrint.Print vbTab; "High School: "; cboHighSchool.Text
    
    frmPrint.Print
    frmPrint.Print vbTab; vbTab; "by Douglas Nielson"


End Sub

Sub PrintStudent()
    Printer.FontSize = 18
    Printer.FontName = "Arial"
    Printer.Print vbTab; "Student Summary"
    Printer.Print vbTab; "Name: "; txtName.Text
    Printer.Print vbTab; "Units Complete: "; txtUnits.Text
    If chkDeanList.Value = vbUnchecked Then
        Printer.Print vbTab; "NOT ";
    Else
        Printer.Print vbTab;
    End If
    Printer.Print "On Dean's List"
    Printer.Print vbTab; "Class: ";
    If optFreshman.Value = True Then
        Printer.Print "Freshman"
    ElseIf optSophomore.Value = True Then
        Printer.Print "Sophomore"
    ElseIf optJunior.Value = True Then
        Printer.Print "Junior"
    Else
        Printer.Print "Senior"
    End If
    Printer.Print vbTab; "Major: ";
    If lstMajor.ListIndex = -1 Then
        Printer.Print "None Selected"
    Else
        Printer.Print lstMajor.List(lstMajor.ListIndex)
    End If
    Printer.Print vbTab; "High School: "; cboHighSchool.Text
    
    Printer.Print
    Printer.Print vbTab; vbTab; "by Douglas Nielson"
    Printer.EndDoc      ' flush the printer buffer
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFilePrintSchool_Click()
    Dim intLoop As Integer
    mstrTypePrint = "HighSchool"
    frmPrint.Show
    Me.Hide
    frmPrint.Cls        ' clear old printing
    
    frmPrint.FontSize = 10
    frmPrint.FontName = "Arial"
    
    frmPrint.Print "High School Listing"

    ' print entire list of high schools
    For intLoop = 0 To cboHighSchool.ListCount - 1
        frmPrint.Print cboHighSchool.List(intLoop)
    Next    ' bottom of For loop
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Student Information Program" & vbCrLf & _
        "by Douglas Nielson", vbInformation + vbOKOnly, _
        "About Information"
End Sub
