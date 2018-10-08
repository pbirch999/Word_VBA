Attribute VB_Name = "ExtractACRONYMSToNewDocument"
'Attribute VB_Name = "basACRONYM_Extract"
Option Explicit

Sub ExtractACRONYMSToNewDocument()

    '=========================
    'Macro created 2008 by Lene Fredborg, DocTools - www.thedoctools.com
    'THIS MACRO IS COPYRIGHT. YOU ARE WELCOME TO USE THE MACRO BUT YOU MUST KEEP THE LINE ABOVE.
    'YOU ARE NOT ALLOWED TO PUBLISH THE MACRO AS YOUR OWN, IN WHOLE OR IN PART.
    '=========================
    'The macro creates a new document,
    'finds all words consisting of 3 or more uppercase letters
    'in the active document and inserts the words
    'in column 1 of a 3-column table in the new document
    'Each acronym is added only once
    'Use column 2 for definitions
    'Page number of first occurrence is added by the macro in column 3
    
    'Minor adjustments are made to the styles used
    'You may need to change the style settings and table layout to fit your needs
    '=========================

    Dim oDoc_Source As Document
    Dim oDoc_Target As Document
    Dim strListSep As String
    Dim strAcronym As String
    Dim oTable As Table
    Dim oRange As Range
    Dim n As Long
    Dim strAllFound As String
    Dim Title As String
    Dim Msg As String

    Title = "Extract Acronyms to New Document"
    
    'Show msg - stop if user does not click Yes
    Msg = "This macro finds all words consisting of 3 or more " & _
        "uppercase letters and extracts the words to a table " & _
        "in a new document where you can add definitions." & vbCr & vbCr & _
        "Do you want to continue?"

    If MsgBox(Msg, vbYesNo + vbQuestion, Title) <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    'Find the list separator from international settings
    'May be a comma or semicolon depending on the country
    strListSep = Application.International(wdListSeparator)
    
    'Start a string to be used for storing names of acronyms found
    strAllFound = "#"
    
    Set oDoc_Source = ActiveDocument
    
    'Create new document for acronyms
    Set oDoc_Target = Documents.Add
    
    With oDoc_Target
        'Make sure document is empty
        .Range = ""
    
        'Insert info in header - change date format as you wish
        .PageSetup.TopMargin = CentimetersToPoints(3)
        .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
            "Acronyms extracted from: " & oDoc_Source.FullName & vbCr & _
            "Created by: " & Application.UserName & vbCr & _
            "Creation date: " & Format(Date, "MMMM d, yyyy")
                
        'Adjust the Normal style and Header style
        With .Styles(wdStyleNormal)
            .Font.Name = "Arial"
            .Font.Size = 10
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.SpaceAfter = 6
        End With
    
        With .Styles(wdStyleHeader)
            .Font.Size = 8
            .ParagraphFormat.SpaceAfter = 0
        End With
        
        'Insert a table with room for acronym and definition
        Set oTable = .Tables.Add(Range:=.Range, NumRows:=2, NumColumns:=3)
        With oTable
            'Format the table a bit
            'Insert headings
            .Range.Style = wdStyleNormal
            .AllowAutoFit = False
            
            .Cell(1, 1).Range.Text = "Acronym"
            .Cell(1, 2).Range.Text = "Definition"
            .Cell(1, 3).Range.Text = "Page"
            'Set row as heading row
            .Rows(1).HeadingFormat = True
            .Rows(1).Range.Font.Bold = True
            .PreferredWidthType = wdPreferredWidthPercent
            .Columns(1).PreferredWidth = 20
            .Columns(2).PreferredWidth = 70
            .Columns(3).PreferredWidth = 10
        End With
    End With
    
    With oDoc_Source
        Set oRange = .Range
        
        n = 1 'used to count below
        
        With oRange.Find
            'Use wildcard search to find strings consisting of 3 or more uppercase letters
            'Set the search conditions
            'NOTE: If you want to find acronyms with e.g. 2 or more letters,
            'change 3 to 2 in the line below
            .Text = "<[A-Z]{3" & strListSep & "}>"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = True
            
            'Perform the search
            Do While .Execute
                'Continue while found
                strAcronym = oRange
                'Insert in target doc
                
                'If strAcronym is already in strAllFound, do not add again
                If InStr(1, strAllFound, "#" & strAcronym & "#") = 0 Then
                    'Add new row in table from second acronym
                    If n > 1 Then oTable.Rows.Add
                    'Was not found before
                    strAllFound = strAllFound & strAcronym & "#"
                    
                    'Insert in column 1 in oTable
                    'Compensate for heading row
                    With oTable
                        .Cell(n + 1, 1).Range.Text = strAcronym
                        'Insert page number in column 3
                        .Cell(n + 1, 3).Range.Text = oRange.Information(wdActiveEndPageNumber)
                    End With
                    
                    n = n + 1
                End If
            Loop
        End With
    End With
    
    'Sort the acronyms alphabetically - skip if only 1 found
    If n > 2 Then
        With Selection
            .Sort ExcludeHeader:=True, FieldNumber:="Column 1", SortFieldType _
                :=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
            
            'Go to start of document
            .HomeKey (wdStory)
        End With
    End If
        
    Application.ScreenUpdating = True
    
    'If no acronyms found, show msg and close new document without saving
    'Else keep open
    If n = 1 Then
        Msg = "No acronyms found."
        oDoc_Target.Close savechanges:=wdDoNotSaveChanges
    Else
        Msg = "Finished extracting " & n - 1 & " acronymn(s) to a new document."
    End If
    
    MsgBox Msg, vbOKOnly, Title
    
    'Clean up
    Set oRange = Nothing
    Set oDoc_Source = Nothing
    Set oDoc_Target = Nothing
    Set oTable = Nothing
    
End Sub



