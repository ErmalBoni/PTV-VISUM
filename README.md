Attribute VB_Name = "Module1"
Option Explicit

' Subroutine triggered by a button click to load a Visum version file
Sub Button1_Click()
    Dim Vis As Visum
    Dim verfile As String
    
    ' Create a new instance of the Visum application
    Set Vis = New Visum
    
    ' Get the file path from a cell in the active sheet (e.g., cell D3)
    ' IMPORTANT: If you want to use a different cell, change the "D3" reference below to the desired cell.
    verfile = Application.ActiveSheet.Range("D3").Text
    
    ' Load the Visum version file
    Vis.LoadVersion verfile
    
    ' Display a message box confirming the version has been loaded
    MsgBox "Version loaded"
End Sub
