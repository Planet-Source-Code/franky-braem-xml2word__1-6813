Attribute VB_Name = "Global"
Option Explicit

Public oWord As Word.Application
Public nRef As Integer

Public Function GetPath(sPath As String) As String

    Dim lPos As Long

    lPos = InStrRev(sPath, "\")
    If lPos > 0 Then
        GetPath = Left(sPath, lPos)
    Else
        GetPath = ""
    End If

End Function
