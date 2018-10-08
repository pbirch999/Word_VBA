Attribute VB_Name = "Module3"
Option Explicit
Sub CaptionFigures()
Dim intCt As Integer
Dim i As Integer

For i = 1 To ActiveDocument.InlineShapes.Count
    If .InlineShape.Item(i).Type = wdInlineShapePicture Then
        .Select
        .InsertCaption Label:="Figure", Title:="Hey", Position:=wdCaptionPositionBelow
        End If
    Next i
End Sub


Sub RefToLibrary()
    ' create a reference to the VBA Extensibility library.
    On Error Resume Next            ' in case the reference already exits
    ThisWorkbook.VBProject.References _
                  .AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 0
End Sub

