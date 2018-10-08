Attribute VB_Name = "SetTableWidth"
Sub SetTableWidth()

Dim t As Table
Dim ccount As Integer
For Each t In ActiveDocument.Tables
    With t
        ccount = .Columns.Count
        If ccount = 7 Then
         Debug.Print t.Title
            .Columns.AutoFit
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
        End If
    End With
Next
End Sub




