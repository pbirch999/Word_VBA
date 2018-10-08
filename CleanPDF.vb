Sub PastePDFClean()

' Works on whatever is on Clipboard. Copy the selection from PDF, switch to Word and run the macro

    Dim MyData As DataObject
    Dim sTextIn As String
    Dim x As Integer
    Dim y As Integer

    Set MyData = New DataObject
    MyData.GetFromClipboard
    sTextIn = MyData.GetText

    x = InStr(sTextIn, vbCr)
    y = 1
    While x > 0
        sTextIn = Left(sTextIn, x - 1) & Mid(sTextIn, x + 1)
        y = x + 1
        x = InStr(y, sTextIn, vbCr)
    Wend

    Selection.TypeText sTextIn
    Set MyData = Nothing
End Sub