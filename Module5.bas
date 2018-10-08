Attribute VB_Name = "Module5"
Option Explicit

Sub DeleteModule()
    'Dim VBProj As VBIDE.Project
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set vbProj = ActiveDocument.VBProject
    Set vbComp = vbProj.VBComponents("Module3")
    vbProj.VBComponents.Remove vbComp

End Sub


Sub RefToLibrary()
    ' create a reference to the VBA Extensibility library.
    On Error Resume Next
    ' in case the reference already exits
    ThisDocument.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 0
   End Sub

Sub RenameModule()
Dim vbProj As VBIDE.VBProject
Dim vbComp As VBIDE.VBComponent
Dim mods As Collection

Set vbProj = ActiveDocument.VBProject

For Each vbComp In vbProj.VBComponents
    If (vbComp.Type = vbext_ct_StdModule) Then
    Debug.Print vbComp.Name
  
   
    End If
Next vbComp

    
    




'ActiveDocument.VBProject.VBComponents("Module3").Name = "Trash"

End Sub
