Option Explicit

Dim objWord, objDoc, fso
Dim filePath, searchText, replaceText, caseSensitive
Dim totalReplacements

totalReplacements = 0

If WScript.Arguments.Count < 3 Then
    WScript.Echo "ERROR: Need at least 3 arguments"
    WScript.Quit 1
End If

filePath = WScript.Arguments(0)
searchText = WScript.Arguments(1)
replaceText = WScript.Arguments(2)

' Check if case sensitivity argument is provided, default to False
If WScript.Arguments.Count >= 4 Then
    caseSensitive = CBool(WScript.Arguments(3))
Else
    caseSensitive = False
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(filePath) Then
    WScript.Echo "ERROR: File not found: " & filePath
    WScript.Quit 1
End If

On Error Resume Next
Set objWord = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot start Word: " & Err.Description
    WScript.Quit 1
End If

objWord.Visible = False
objWord.DisplayAlerts = False
objWord.ScreenUpdating = False

Dim fullPath
fullPath = fso.GetAbsolutePathName(filePath)

Set objDoc = objWord.Documents.Open(fullPath)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot open document"
    objWord.ScreenUpdating = True
    objWord.Quit
    WScript.Quit 1
End If

WScript.Echo "SUCCESS: Document opened"

' FORMATTING-PRESERVING: Replace in shapes using Find & Replace
Dim shape
Dim shapeCount
shapeCount = 0

For Each shape In objDoc.Shapes
    On Error Resume Next
    If shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            If Len(shape.TextFrame.TextRange.Text) > 1 Then
                Dim shapeRange
                Set shapeRange = shape.TextFrame.TextRange
                
                ' Count actual occurrences BEFORE replacement
                Dim shapeReplacements
                shapeReplacements = CountOccurrences(shapeRange.Text, searchText, caseSensitive)
                
                If shapeReplacements > 0 Then
                    ' Use Word's Find & Replace to preserve formatting
                    With shapeRange.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = searchText
                        .Replacement.Text = replaceText
                        .Forward = True
                        .Wrap = 1
                        .Format = False
                        .MatchCase = caseSensitive
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        
                        ' Execute replace all and add actual count
                        .Execute , , , , , , , , , , 2
                        shapeCount = shapeCount + shapeReplacements
                    End With
                End If
            End If
        End If
    End If
    On Error GoTo 0
Next

If shapeCount > 0 Then
    WScript.Echo "SHAPES: " & shapeCount & " replacements"
    totalReplacements = totalReplacements + shapeCount
Else
    WScript.Echo "SHAPES: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in headers
Dim section, headerCount, i
headerCount = 0
For Each section In objDoc.Sections
    For i = 1 To 3
        On Error Resume Next
        If section.Headers(i).Exists Then
            If Len(section.Headers(i).Range.Text) > 1 Then
                Dim headerRange
                Set headerRange = section.Headers(i).Range
                
                ' Count actual occurrences BEFORE replacement
                Dim headerReplacements
                headerReplacements = CountOccurrences(headerRange.Text, searchText, caseSensitive)
                
                If headerReplacements > 0 Then
                    ' Use Word's Find & Replace to preserve formatting
                    With headerRange.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = searchText
                        .Replacement.Text = replaceText
                        .Forward = True
                        .Wrap = 1
                        .Format = False
                        .MatchCase = caseSensitive
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        
                        ' Execute replace all and add actual count
                        .Execute , , , , , , , , , , 2
                        headerCount = headerCount + headerReplacements
                    End With
                End If
            End If
        End If
        On Error GoTo 0
    Next
Next

If headerCount > 0 Then
    WScript.Echo "HEADERS: " & headerCount & " replacements"
    totalReplacements = totalReplacements + headerCount
Else
    WScript.Echo "HEADERS: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in footers
Dim footerCount
footerCount = 0
For Each section In objDoc.Sections
    For i = 1 To 3
        On Error Resume Next
        If section.Footers(i).Exists Then
            If Len(section.Footers(i).Range.Text) > 1 Then
                Dim footerRange
                Set footerRange = section.Footers(i).Range
                
                ' Count actual occurrences BEFORE replacement
                Dim footerReplacements
                footerReplacements = CountOccurrences(footerRange.Text, searchText, caseSensitive)
                
                If footerReplacements > 0 Then
                    ' Use Word's Find & Replace to preserve formatting
                    With footerRange.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = searchText
                        .Replacement.Text = replaceText
                        .Forward = True
                        .Wrap = 1
                        .Format = False
                        .MatchCase = caseSensitive
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        
                        ' Execute replace all and add actual count
                        .Execute , , , , , , , , , , 2
                        footerCount = footerCount + footerReplacements
                    End With
                End If
            End If
        End If
        On Error GoTo 0
    Next
Next

If footerCount > 0 Then
    WScript.Echo "FOOTERS: " & footerCount & " replacements"
    totalReplacements = totalReplacements + footerCount
Else
    WScript.Echo "FOOTERS: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in footnotes
Dim footnote, footnoteCount
footnoteCount = 0
On Error Resume Next
For Each footnote In objDoc.Footnotes
    If Len(footnote.Range.Text) > 1 Then
        Dim footnoteRange
        Set footnoteRange = footnote.Range
        
        ' Count actual occurrences BEFORE replacement
        Dim footnoteReplacements
        footnoteReplacements = CountOccurrences(footnoteRange.Text, searchText, caseSensitive)
        
        If footnoteReplacements > 0 Then
            With footnoteRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = searchText
                .Replacement.Text = replaceText
                .Forward = True
                .Wrap = 1
                .Format = False
                .MatchCase = caseSensitive
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                .Execute , , , , , , , , , , 2
                footnoteCount = footnoteCount + footnoteReplacements
            End With
        End If
    End If
Next
On Error GoTo 0

If footnoteCount > 0 Then
    WScript.Echo "FOOTNOTES: " & footnoteCount & " replacements"
    totalReplacements = totalReplacements + footnoteCount
Else
    WScript.Echo "FOOTNOTES: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in endnotes
Dim endnote, endnoteCount
endnoteCount = 0
On Error Resume Next
For Each endnote In objDoc.Endnotes
    If Len(endnote.Range.Text) > 1 Then
        Dim endnoteRange
        Set endnoteRange = endnote.Range
        
        ' Count actual occurrences BEFORE replacement
        Dim endnoteReplacements
        endnoteReplacements = CountOccurrences(endnoteRange.Text, searchText, caseSensitive)
        
        If endnoteReplacements > 0 Then
            With endnoteRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = searchText
                .Replacement.Text = replaceText
                .Forward = True
                .Wrap = 1
                .Format = False
                .MatchCase = caseSensitive
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                .Execute , , , , , , , , , , 2
                endnoteCount = endnoteCount + endnoteReplacements
            End With
        End If
    End If
Next
On Error GoTo 0

If endnoteCount > 0 Then
    WScript.Echo "ENDNOTES: " & endnoteCount & " replacements"
    totalReplacements = totalReplacements + endnoteCount
Else
    WScript.Echo "ENDNOTES: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in form fields
Dim field, formFieldCount
formFieldCount = 0
On Error Resume Next
For Each field In objDoc.FormFields
    If field.Type = 70 Then ' wdFieldFormTextInput
        Dim originalText, newText
        originalText = field.Result
        
        If Len(originalText) > 0 Then
            ' Count actual occurrences BEFORE replacement
            Dim formFieldReplacements
            formFieldReplacements = CountOccurrences(originalText, searchText, caseSensitive)
            
            If formFieldReplacements > 0 Then
                If caseSensitive Then
                    newText = Replace(originalText, searchText, replaceText)
                Else
                    newText = Replace(originalText, searchText, replaceText, 1, -1, 1) ' vbTextCompare
                End If
                field.Result = newText
                formFieldCount = formFieldCount + formFieldReplacements
            End If
        End If
    End If
Next
On Error GoTo 0

If formFieldCount > 0 Then
    WScript.Echo "FORMFIELDS: " & formFieldCount & " replacements"
    totalReplacements = totalReplacements + formFieldCount
Else
    WScript.Echo "FORMFIELDS: 0 replacements"
End If

' SAFER: Replace in hyperlinks (ONLY those NOT in shapes to avoid double counting)
Dim hyperlink, hyperlinkCount
hyperlinkCount = 0
On Error Resume Next
For Each hyperlink In objDoc.Hyperlinks
    If Len(hyperlink.TextToDisplay) > 0 Then
        ' Check if this hyperlink is within a shape
        Dim isInShape
        isInShape = False
        
        ' Check if hyperlink range is within any shape
        For Each shape In objDoc.Shapes
            On Error Resume Next
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    ' Simple check: if hyperlink start is within shape range
                    If hyperlink.Range.Start >= shape.TextFrame.TextRange.Start And hyperlink.Range.End <= shape.TextFrame.TextRange.End Then
                        isInShape = True
                        Exit For
                    End If
                End If
            End If
            On Error GoTo 0
        Next
        
        ' Only process if NOT in shape
        If Not isInShape Then
            ' Count actual occurrences in display text
            Dim hyperlinkReplacements
            hyperlinkReplacements = CountOccurrences(hyperlink.TextToDisplay, searchText, caseSensitive)
            
            If hyperlinkReplacements > 0 Then
                ' SAFER: Replace only the display text, preserve the link
                Dim newDisplayText
                If caseSensitive Then
                    newDisplayText = Replace(hyperlink.TextToDisplay, searchText, replaceText)
                Else
                    newDisplayText = Replace(hyperlink.TextToDisplay, searchText, replaceText, 1, -1, 1)
                End If
                
                ' Update display text while preserving hyperlink
                hyperlink.TextToDisplay = newDisplayText
                hyperlinkCount = hyperlinkCount + hyperlinkReplacements
            End If
        End If
    End If
Next
On Error GoTo 0

If hyperlinkCount > 0 Then
    WScript.Echo "HYPERLINKS: " & hyperlinkCount & " replacements"
    totalReplacements = totalReplacements + hyperlinkCount
Else
    WScript.Echo "HYPERLINKS: 0 replacements"
End If

' FORMATTING-PRESERVING: Replace in main document content (paragraphs + tables)
Dim mainContentCount
mainContentCount = 0

' Count actual occurrences in main content BEFORE replacement
Dim mainContentReplacements
mainContentReplacements = CountOccurrences(objDoc.Content.Text, searchText, caseSensitive)

If mainContentReplacements > 0 Then
    ' Use Word's Find & Replace on main document content to preserve formatting
    With objDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = searchText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = caseSensitive
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Execute replace all and add actual count
        .Execute , , , , , , , , , , 2
        mainContentCount = mainContentCount + mainContentReplacements
    End With
End If

If mainContentCount > 0 Then
    WScript.Echo "MAINCONTENT: " & mainContentCount & " replacements"
    totalReplacements = totalReplacements + mainContentCount
Else
    WScript.Echo "MAINCONTENT: 0 replacements"
End If


objWord.ScreenUpdating = True

objDoc.Save
objDoc.Close
objWord.Quit

WScript.Echo "SUCCESS: Document saved"
WScript.Echo "RESULT: " & totalReplacements & " total replacements"

WScript.Quit 0

Function CountOccurrences(text, searchFor, caseSensitive)
    If Len(searchFor) = 0 Then
        CountOccurrences = 0
        Exit Function
    End If
    
    If caseSensitive Then
        CountOccurrences = (Len(text) - Len(Replace(text, searchFor, ""))) / Len(searchFor)
    Else
        CountOccurrences = (Len(text) - Len(Replace(LCase(text), LCase(searchFor), ""))) / Len(searchFor)
    End If
End Function
