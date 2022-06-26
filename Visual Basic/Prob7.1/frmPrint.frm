VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print Preview"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Send to Printer"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Print Preview Form

Private Sub cmdExit_Click()
    frmStudent.Show     ' other form is active
    Me.Hide             ' make invisible
    Unload frmPrint     ' remove from memory
End Sub

Private Sub cmdPrint_Click()
    Select Case frmStudent.mstrTypePrint
    Case "Student"
        Printer.FontSize = 18
        Printer.FontName = "Arial"
        Printer.Print vbTab; "Student Summary"
        Printer.Print vbTab; "Name: "; frmStudent.txtName.Text
        Printer.Print vbTab; "Units Complete: "; frmStudent.txtUnits.Text
        If frmStudent.chkDeanList.Value = vbUnchecked Then
            Printer.Print vbTab; "NOT ";
        Else
            Printer.Print vbTab;
        End If
        Printer.Print "On Dean's List"
        Printer.Print vbTab; "Class: ";
        If frmStudent.optFreshman.Value = True Then
            Printer.Print "Freshman"
        ElseIf frmStudent.optSophomore.Value = True Then
            Printer.Print "Sophomore"
        ElseIf frmStudent.optJunior.Value = True Then
            Printer.Print "Junior"
        Else
            Printer.Print "Senior"
        End If
        Printer.Print vbTab; "Major: ";
        If frmStudent.lstMajor.ListIndex = -1 Then
            Printer.Print "None Selected"
        Else
            Printer.Print frmStudent.lstMajor.List(frmStudent.lstMajor.ListIndex)
        End If
        Printer.Print vbTab; "High School: "; frmStudent.cboHighSchool.Text
        
        Printer.Print
        Printer.FontSize = 8
        Printer.Print vbTab; vbTab; "Program by Douglas Nielson"
        Printer.EndDoc      ' flush the printer buffer
    Case "HighSchool"
        Dim intLoop As Integer
        Printer.FontSize = 12
        Printer.FontName = "Arial"
    
        Printer.Print vbTab; "High School Listing"
        ' print entire list of high schools
        For intLoop = 0 To frmStudent.cboHighSchool.ListCount - 1
            Printer.Print vbTab; frmStudent.cboHighSchool.List(intLoop)
        Next    ' bottom of For loop
        Printer.EndDoc      ' flush the printer buffer
    End Select
End Sub
