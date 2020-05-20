Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : RibbonItemLaunch
' Author    : Adam Kauffman
' Date      : 2020-05-20
' Purpose   : Callback for loadGUI onAction
'---------------------------------------------------------------------------------------
'
Public Sub RibbonItemLaunch(control As IRibbonControl)
    AddInMenuItemLaunch
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RibbonItemLaunch
' Author    : Adam Kauffman
' Date      : 2020-05-20
' Purpose   : Callback for exportSrc onAction
'---------------------------------------------------------------------------------------
'
Public Sub RibbonItemExport(control As IRibbonControl)
    AddInMenuItemExport
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RibbonItemImport
' Author    : Adam Kauffman
' Date      : 2020-05-20
' Purpose   : Callback for importSrc onAction
'---------------------------------------------------------------------------------------
'
Public Sub RibbonItemImport(control As IRibbonControl)
    Form_frmVCSMain.Visible = True
    DoEvents
    Form_frmVCSMain.cmdBuild_Click
End Sub


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