Attribute VB_Name = "NewMacros"
Option Explicit


Sub TC_Update_Fields()
Attribute TC_Update_Fields.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.TC_Update_Fields"
'
' TC_Update_Fields Macro
'
'
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 23")).Select
  '  ActiveDocument.Shapes.Range(Array("Text Box 23")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 23")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    Selection.Fields.Update
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select 'version number
    Selection.Fields.Update
End Sub
Sub AutoOpen()
Attribute AutoOpen.VB_Description = "Adjust zoom level "
Attribute AutoOpen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.AutoOpen"
'
' AutoOpen Macro
' Adjust zoom level
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = 125

End Sub
