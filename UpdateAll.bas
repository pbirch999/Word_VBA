Attribute VB_Name = "Module1"
Option Explicit

Sub UpdateALL()
Dim oStory As Range
For Each oStory In ActiveDocument.StoryRanges
    oStory.Fields.Update
If oStory.StoryType <> wdMainTextStory Then
    While Not (oStory.NextStoryRange Is Nothing)
        Set oStory = oStory.NextStoryRange
        oStory.Fields.Update
    Wend
End If
Next oStory
lbl_Exit:
Set oStory = Nothing
    Exit Sub
End Sub
