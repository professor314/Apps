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
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
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
' Example of menu building chapter 5
' Douglas Nielson
Option Explicit

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFilePrint_Click()
    MsgBox "No print for now", vbOKOnly, "Print Message"
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Progam by Doublas Nielson" & vbCrLf & _
        "CIS 160 Visual Basic", vbOKOnly, _
        "About Message"
End Sub
