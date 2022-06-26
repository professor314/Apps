Attribute VB_Name = "StartModule"
' startup module for program
Option Explicit
Sub Main()
    ' code from book page 227
    frmSplash.Show vbModeless   ' start splash screen
    ' load brings form into memory but it is not visible
    Load frmWork                ' load second form
End Sub
