VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Sample of Menus"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditFont 
         Caption         =   "&Font"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Example of menu building
' Chapter 5
' Sean Connolly
Option Explicit

Private Sub mnuHelpAbout_Click()
    MsgBox "Program by" & vbCrLf & "Sean Connolly", vbOKOnly, "About"
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFilePrint_Click()
    MsgBox "No Printing still!", vbOKOnly, "Error"
End Sub
