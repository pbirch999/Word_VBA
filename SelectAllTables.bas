Attribute VB_Name = "SelectAllTables"
'Select all tables in a Word document.
Sub SelectAllTables()
  Dim objDoc As Document
  Dim objTable As Table
 
  Application.ScreenUpdating = False
 
  'Initialization
  Set objDoc = ActiveDocument
 
  'Set each table in document as a range editable to everyone.
  With objDoc
  For Each objTable In .Tables
    objTable.Range.Editors.Add wdEditorEveryone
  Next
  objDoc.SelectAllEditableRanges wdEditorEveryone
  objDoc.DeleteAllEditableRanges wdEditorEveryone
  Application.ScreenUpdating = True
  End With
End Sub
Sub changeSpacing()
'
' changeSpacing Macro
'
'
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 3
        .SpaceBeforeAuto = False
        .SpaceAfter = 3
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
End Sub

