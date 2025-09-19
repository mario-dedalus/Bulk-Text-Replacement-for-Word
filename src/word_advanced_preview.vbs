' Word Advanced Preview Script - READ ONLY - COMPREHENSIVE VERSION
Dim args, filePath, searchText, caseSensitive
Set args = WScript.Arguments

If args.Count < 3 Then
    WScript.Echo "Usage: script.vbs <filepath> <searchtext> <casesensitive>"
    WScript.Quit 1
End If

filePath = args(0)
searchText = args(1)
caseSensitive = CBool(args(2))

Dim wordApp, doc
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = False
wordApp.ScreenUpdating = False

On Error Resume Next
Set doc = wordApp.Documents.Open(filePath, , True) ' Read-only
If Err.Number <> 0 Then
    WScript.Echo "Error opening document"
    wordApp.ScreenUpdating = True
    wordApp.Quit
    WScript.Quit 1
End If

Dim shapesCount, headersCount, footersCount, footnotesCount, endnotesCount, formFieldsCount, hyperlinksCount
shapesCount = 0
headersCount = 0
footersCount = 0
footnotesCount = 0
endnotesCount = 0
formFieldsCount = 0
hyperlinksCount = 0

' Count in shapes
For Each shape In doc.Shapes
    On Error Resume Next
    If shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            Dim shapeText
            shapeText = shape.TextFrame.TextRange.Text
            If Len(shapeText) > 1 Then
                If caseSensitive Then
                    shapesCount = shapesCount + CountOccurrences(shapeText, searchText, True)
                Else
                    shapesCount = shapesCount + CountOccurrences(shapeText, searchText, False)
                End If
            End If
        End If
    End If
    On Error GoTo 0
Next

' Count in headers and footers
For Each section In doc.Sections
    ' Headers
    For i = 1 To 3
        On Error Resume Next
        If section.Headers(i).Exists Then
            Dim headerText
            headerText = section.Headers(i).Range.Text
            If Len(headerText) > 1 Then
                If caseSensitive Then
                    headersCount = headersCount + CountOccurrences(headerText, searchText, True)
                Else
                    headersCount = headersCount + CountOccurrences(headerText, searchText, False)
                End If
            End If
        End If
        On Error GoTo 0
    Next
    
    ' Footers
    For i = 1 To 3
        On Error Resume Next
        If section.Footers(i).Exists Then
            Dim footerText
            footerText = section.Footers(i).Range.Text
            If Len(footerText) > 1 Then
                If caseSensitive Then
                    footersCount = footersCount + CountOccurrences(footerText, searchText, True)
                Else
                    footersCount = footersCount + CountOccurrences(footerText, searchText, False)
                End If
            End If
        End If
        On Error GoTo 0
    Next
Next

' Count in footnotes
On Error Resume Next
For Each footnote In doc.Footnotes
    Dim footnoteText
    footnoteText = footnote.Range.Text
    If Len(footnoteText) > 1 Then
        If caseSensitive Then
            footnotesCount = footnotesCount + CountOccurrences(footnoteText, searchText, True)
        Else
            footnotesCount = footnotesCount + CountOccurrences(footnoteText, searchText, False)
        End If
    End If
Next
On Error GoTo 0

' Count in endnotes
On Error Resume Next
For Each endnote In doc.Endnotes
    Dim endnoteText
    endnoteText = endnote.Range.Text
    If Len(endnoteText) > 1 Then
        If caseSensitive Then
            endnotesCount = endnotesCount + CountOccurrences(endnoteText, searchText, True)
        Else
            endnotesCount = endnotesCount + CountOccurrences(endnoteText, searchText, False)
        End If
    End If
Next
On Error GoTo 0

' Count in form fields
On Error Resume Next
For Each field In doc.FormFields
    If field.Type = 70 Then ' wdFieldFormTextInput
        Dim fieldText
        fieldText = field.Result
        If Len(fieldText) > 0 Then
            If caseSensitive Then
                formFieldsCount = formFieldsCount + CountOccurrences(fieldText, searchText, True)
            Else
                formFieldsCount = formFieldsCount + CountOccurrences(fieldText, searchText, False)
            End If
        End If
    End If
Next
On Error GoTo 0

' Count in hyperlinks (ONLY those NOT in shapes to avoid double counting)
On Error Resume Next
For Each hyperlink In doc.Hyperlinks
    Dim hyperlinkText
    hyperlinkText = hyperlink.TextToDisplay
    If Len(hyperlinkText) > 0 Then
        ' Check if this hyperlink is within a shape
        Dim isInShape
        isInShape = False
        
        ' Check if hyperlink range is within any shape
        For Each shape In doc.Shapes
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
        
        ' Only count if NOT in shape
        If Not isInShape Then
            If caseSensitive Then
                hyperlinksCount = hyperlinksCount + CountOccurrences(hyperlinkText, searchText, True)
            Else
                hyperlinksCount = hyperlinksCount + CountOccurrences(hyperlinkText, searchText, False)
            End If
        End If
    End If
Next
On Error GoTo 0

WScript.Echo "SHAPES: " & shapesCount & " matches"
WScript.Echo "HEADERS: " & headersCount & " matches"
WScript.Echo "FOOTERS: " & footersCount & " matches"
WScript.Echo "FOOTNOTES: " & footnotesCount & " matches"
WScript.Echo "ENDNOTES: " & endnotesCount & " matches"
WScript.Echo "FORMFIELDS: " & formFieldsCount & " matches"
WScript.Echo "HYPERLINKS: " & hyperlinksCount & " matches"

' Count in main document content (paragraphs + tables)
Dim mainContentCount
mainContentCount = 0
On Error Resume Next
Dim mainContentText
mainContentText = doc.Content.Text
If Len(mainContentText) > 1 Then
    If caseSensitive Then
        mainContentCount = CountOccurrences(mainContentText, searchText, True)
    Else
        mainContentCount = CountOccurrences(mainContentText, searchText, False)
    End If
End If
On Error GoTo 0

WScript.Echo "MAINCONTENT: " & mainContentCount & " matches"

wordApp.ScreenUpdating = True
doc.Close False
wordApp.Quit

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
