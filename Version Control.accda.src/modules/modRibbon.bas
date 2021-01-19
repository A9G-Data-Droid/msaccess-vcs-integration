Option Compare Database
Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : LoadRibbons
' Author    : Adam Kauffman
' Date      : 2020-05-20
' Purpose   : Load all the ribbons defined in this database
'---------------------------------------------------------------------------------------
'
Public Sub LoadRibbons()
    On Error GoTo ErrorHandler
    With CurrentDb.OpenRecordset("USysRibbons")
        .MoveFirst
        While Not .EOF
            Application.LoadCustomUI .Fields("RibbonName").Value, .Fields("RibbonXml").Value
            .MoveNext
        Wend
        
        .Close
    End With

ErrorHandler:
    If Err.Number = 32610 Then
        Resume Next
    ElseIf Err.Number > 0 Then
        Err.Raise 10016, "modRibbon.LoadRibbons", "Error loading ribbons: " & Err.Number & " - " & Err.Description
    End If
End Sub