Attribute VB_Name = "modStart"
Option Explicit

Sub Main()
    On Error Resume Next
    
    If App.PrevInstance Then
        MsgBox "The App is already started!", vbCritical, "Error"
        Exit Sub
    End If
    
    If Command = vbNullString Then
        frmMain.Show
    Else
        If FileExist(Command) Then
            frmMain.picMain.Picture = LoadPicture(Command)
            Call ResizePicBoxes
            frmMain.Show
        End If
    End If
    
End Sub
