Attribute VB_Name = "Module1"
Sub multisheet()
    Dim sheets As Worksheet
        Application.ScreenUpdating = False
            For Each sheets In Worksheets
            sheets.Select
            Call pleasework
        Next
        Application.ScreenUpdating = True
End Sub




