' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009-2012 Mendeley Ltd.
'
' Licensed under the Educational Community License, Version 1.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.opensource.org/licenses/ecl1.php
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' ***** END LICENSE BLOCK *****

' author: steve.ridout@mendeley.com

' Note: The UUIDs in these tests refer to documents in the tests/testDatabase@test.com@local.sqlite
' Mendeley Desktop database. See the README.txt for instructions to run the tests.

Option Explicit

Function testsPath() As String
    ' Read path from environment variable
    testsPath = Environ("MENDELEY_OO_TEST_FILES")
    If testsPath = "" Then
        MsgBox "Please ensure the MENDELEY_OO_TEST_FILES environment variable is set to " & Chr(13) &_
               "the directory containing your *.odt test files before running the tests."
    End If
End Function

Function outputPath() As String
    outputPath = testsPath() & "/output"
End Function

Function crLf() As String
    crLf = Chr(13) & Chr(10)
End Function

Sub runUnitTests()
    unitTest = True
    
    Call MkDir(outputPath())
    
Call testApiCallMultipleArgs
    Call testInsertCitation
    Call testEditCitation
    Call testMergeCitations
    Call testApplyFormatting
    Call testRefreshDocument
    Call testChangeCitationStyle
    Call testInsertBibliography
    Call testExportOpenOfficeSimple

    unitTest = False
End Sub

' -- utility functions --

Sub exportExpected(documentName)
    ' Export expected txt file       
    dim exportProperties(1) as new com.sun.star.beans.PropertyValue
    exportProperties(0).Name = "FilterName"
    exportProperties(0).Value = "Text"
    ThisComponent.storeToUrl("file:///" & outputPath() & documentName & "-expected.txt", exportProperties())
End Sub

Sub exportActual(documentName)
    ' Export expected txt file       
    dim exportProperties(1) as new com.sun.star.beans.PropertyValue
    exportProperties(0).Name = "FilterName"
    exportProperties(0).Value = "Text"
    ThisComponent.storeToUrl("file:///" & outputPath() & documentName & "-actual.txt", exportProperties())
End Sub

Sub exportAsOdt(documentName)
    ' Export expected txt file       
    dim exportProperties(1) as new com.sun.star.beans.PropertyValue
    exportProperties(0).Name = "FilterName"
    exportProperties(0).Value = "writer8"
    ThisComponent.storeToUrl("file:///" & outputPath() & documentName & ".odt", exportProperties())
End Sub

Sub appendText(textToAppend)
    Call thisComponent.getText().insertString(thisComponent.getText().end(), textToAppend, false)
End Sub

Function newDocument()
    newDocument = StarDesktop.loadComponentFromURL( "private:factory/swriter", "_blank", 0, Array() )
    Call setCitationStyle(DEFAULT_CITATION_STYLE)
End Function

Function documentText() As String
    Dim result As String
    result = thisComponent.getText().getString()
    
    documentText = result
End Function

Function normaliseString(inputString As String) As String
    Dim result
    ' Strip out all instances of Chr(8288)
    ' these are zero width non-breaking spaces
    ' if it turns out they are important we can remove this and add them to the expected strings
    result = Replace(inputString, Chr(8288), "")

    ' normalise line endings to LF
    result = Replace(result, crLf(), Chr(10))
    result = Replace(result, Chr(13), Chr(10))
        
    normaliseString = result
End Function

Function compareStrings(actual As String, expected As String, location As String) As Boolean
    compareStrings = True
    
    Dim normalisedActual As String
    Dim normalisedExpected As String
    
    normalisedActual = normaliseString(actual)
    normalisedExpected = normaliseString(expected)
    
    If normalisedActual <> normalisedExpected Then
        MsgBox "Test failed at " & location & Chr(13) &_
                "Len(actual):   " & Len(normalisedActual) & Chr(13) &_
                "Len(expected):   " & Len(normalisedExpected) & Chr(13) &_
               "Actual:   " & normalisedActual & Chr(13) &_
               "Expected: " & normalisedExpected
               
        compareStrings = False
        Exit Function
    End If
End Function

' -- tests --

Sub testApiCallMultipleArgs()
    Dim arguments(1 to 2) As String
    arguments(1) = "arg\n 1"
    arguments(2) = "arg 2"
    Dim result As String
    result = mendeleyApiCall("concatenateStringsTest", arguments)
    Call compareStrings(result, "arg\n 1arg 2", "testApiCallMultipleArgs")
End Sub

Sub testRefreshDocument()
    Dim documentName As String
    Dim outputDocumentName As String
    
    ' search for all .docx files
    Dim filename
    filename = Dir(testsPath() & "/*.odt")
    
    Do While filename <> ""
        documentName = Left(filename, Len(filename) - Len(".odt"))
        outputDocumentName = "refreshDocument/" & documentName
        
        Dim url
        url = ConvertToUrl(testsPath & "/" & documentName & ".odt")
        Dim noArgs() 'An empty array for the arguments
        Dim doc
        doc = StarDesktop.LoadComponentFromUrl(url, "_blank", 0, Array())
        
        ' Export expected txt file
        Dim expectedString
        expectedString = documentText()
        'Call exportExpected("refreshDocument/" & outputDocumentName)
        
        ' refresh and export actual xml
        If Not refreshDocument(False) Then
            thisComponent.Text.setString("refreshDocument() failed")
        End If
        
        Dim actualString
        actualString = documentText()
        Call compareStrings(actualString, expectedString, "refreshDocument: " & documentName)
        Call ThisComponent.close(false)
        
        filename = Dir
    Loop
End Sub

Sub testExportOpenOfficeSimple()
    Dim documentName As String

    Call newDocument()
    
    ' Add a citation and bibliogrphy
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    appendText (" some more test" & Chr(13) & "New paragraph." & Chr(13))
    
'    Call ActiveDocument.Select
'    Call Selection.Collapse(wdCollapseEnd)
    Call insertBibliography
    
    ' Export as expected text file
    Dim expectedString As String
    expectedString = documentText()
        
    Dim exportedFileUrl
    exportedFileUrl = ConvertToUrl(outputPath() & "/exportOO-exported.doc")
    
    ' output as oo compatible
    Call exportAsBookmarks(exportedFileUrl)
    Call thisComponent.close(false)
    
    Dim noArgs() 'An empty array for the arguments
    StarDesktop.LoadComponentFromUrl(exportedFileUrl, "_blank", 0, Array())

    ' TODO: The exported file has an extra newline at the end.
    '       Fix issue, then remove the following line:
    expectedString = expectedString & Chr(10)

    Call compareStrings(documentText(), expectedString, "export-OO")
    Call thisComponent.close(false)
End Sub

Sub testInsertCitation()
    Call newDocument()

    ' Add a citation at start
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    Call appendText (" some more test" & Chr(13) & "New paragraph.")
    
    Call privateInsertCitation("{ac45152c-4707-4d3c-928d-2cc59aa386fa}")
    
    compareStrings(documentText(),_
        "(The Mendeley Support Team, 2011) some more test" & crLf() &_
        "New paragraph.(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.)", "testInsertCitation")
    Call thisComponent.close(false)
End Sub

Sub testChangeCitationStyle()
    Call newDocument()
    
    ' Add a citation at start
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    appendText (" ")
    Call privateInsertCitation("{ac45152c-4707-4d3c-928d-2cc59aa386fa}")
    
    Call setCitationStyle("http://www.zotero.org/styles/apa")
    Call refreshDocument(False)
    Call compareStrings(documentText(), "(The Mendeley Support Team, 2011) (Chumbe, Macleod, Barker, Moffat, & Rist, n.d.)", "changeCitation-APA")
    
    Call setCitationStyle("http://www.zotero.org/styles/ieee")
    Call refreshDocument(False)
    
    Call compareStrings(documentText(), "[1] [2]", "changeCitation-IEEE")
    
    Call thisComponent.close(false)
End Sub

Sub testEditCitation()
    Call newDocument()
    
    ' Add a citation at start
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    Call compareStrings(documentText(), "(The Mendeley Support Team, 2011)", "editCitation1")
    
    ' Place cursor within the new citation
    Dim textCursor
    textCursor = thisComponent.getText().createTextCursor()
    Call textCursor.collapseToStart()
    Call textCursor.goRight(1, false)
    thisComponent.getCurrentController().select(textCursor)
    
    ' Add new citation within first one (will act as an edit and replace the first one)
    Call privateInsertCitation("{ac45152c-4707-4d3c-928d-2cc59aa386fa}")
    Call compareStrings(documentText(), "(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.)", "editCitation2")
    Call thisComponent.close(false)
End Sub

Sub testMergeCitations()
    Call newDocument()
    
    ' Merge two citations with a single space gap
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    Call appendText(" ")
    Call privateInsertCitation("{ac45152c-4707-4d3c-928d-2cc59aa386fa}")
        
    ' select whole document
    thisComponent.getCurrentController().select(thisComponent.getText())
    Call mergeCitations
    Call compareStrings(documentText(), "(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.; The Mendeley Support Team, 2011)", "mergeCitationsSingleGap")
    
    ' Merge with no gap
    thisComponent.getCurrentController().select(thisComponent.getText().getEnd())
    Call privateInsertCitation("{d98145e4-d617-4370-b157-08e4ad40151d}")
    
    thisComponent.getCurrentController().select(thisComponent.getText())
    Call mergeCitations
    
    Call compareStrings(documentText(), "(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.; Hu, Chinenov, Kerppola, Hughes, & Arbor, 2002; The Mendeley Support Team, 2011)", "mergeCitationsNoGap")
    
    ' Merge with multiple characters gap
    Call appendText ("gap")
    thisComponent.getCurrentController().select(thisComponent.getText().getEnd())
    Call privateInsertCitation("{d8695892-e5ce-4731-8e42-5921db85182b}")
    
    thisComponent.getCurrentController().select(thisComponent.getText())
    Call mergeCitations
    
    Call compareStrings(documentText(), "(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.; Devbhandari et al., 2007; Hu, Chinenov, Kerppola, Hughes, & Arbor, 2002; The Mendeley Support Team, 2011)", "mergeCitationsMultileGap")
    
    ' Merge four citations
    thisComponent.getText().setString("gap")
    'thisComponent.getCurrentController().select(thisComponent.getText().getEnd())
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    Call appendText ("gap")
    'thisComponent.getCurrentController().select(thisComponent.getText().getEnd())
    Call privateInsertCitation("{ac45152c-4707-4d3c-928d-2cc59aa386fa}")
    Call appendText ("gap")
    'thisComponent.getCurrentController().select(thisComponent.getText().getEnd())
    Call privateInsertCitation("{d98145e4-d617-4370-b157-08e4ad40151d}")
    Call appendText ("gap")
    Call privateInsertCitation("{d8695892-e5ce-4731-8e42-5921db85182b}")
    Call appendText ("gap")
    thisComponent.getCurrentController().select(thisComponent.getText())
    
    Call mergeCitations
    Call compareStrings(documentText(), "gap(Chumbe, Macleod, Barker, Moffat, & Rist, n.d.; Devbhandari et al., 2007; Hu, Chinenov, Kerppola, Hughes, & Arbor, 2002; The Mendeley Support Team, 2011)gap", "mergeCitationsFourAtOnce")
    
    Call thisComponent.close(false)
End Sub

Sub testInsertBibliography()
    Call newDocument()
    
    ' Add a citation at start
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    appendText (" some more test" & Chr(13) & "New paragraph." & Chr(13))
    ' removed: select end
    Call insertBibliography
    
    Call compareStrings(documentText(), "(The Mendeley Support Team, 2011) some more test" + crLf() +_
        "New paragraph." + crLf() +_
        "The Mendeley Support Team. (2011). Getting Started with Mendeley. Mendeley Desktop. London: Mendeley Ltd. Retrieved from http://www.mendeley.com" + crLf() + _
        "", "insertBibliography")
    Call thisComponent.close(false)
End Sub

Sub testApplyFormatting()
    Dim mark 'As field
    Dim rangeToAdd
    
    Call newDocument()
    
    ' no spaces before and after
    Call appendText ("start")
    mark = fnAddMark(thisComponent.getText().getEnd(), "test field 2")
    Call appendText ("end")
    
    Call applyFormatting("<unicode>0061</unicode>", mark)
    
    ' spaces before and after mark
    Call appendText (Chr(13) & "start ")
    Set mark = fnAddMark(thisComponent.getText().getEnd(), "test field")
    Call appendText (" end")
    
    Call applyFormatting("<unicode>0061</unicode>", mark)
    
    ' spaces before and after mark and within mark
    Call appendText (Chr(13) & "start ")
    Set mark = fnAddMark(thisComponent.getText().getEnd(), "test field")
    Call appendText (" end")
    
    Call applyFormatting(" <unicode>0061</unicode> ", mark)
    
    Call appendText (Chr(13) & "start ")
    Set mark = fnAddMark(thisComponent.getText().getEnd(), "test field")
    Call appendText (" end")
    
    Call applyFormatting(" <b>bold</b> ", mark)
    Call compareStrings(documentText(), "start=end" & crLf() & "start = end" & crLf() & _
        "start  =  end" & crLf() & "start  bold  end", "applyFormatting")
    Call thisComponent.close(false)
End Sub

Sub testChangeMarkFormat()
    ' This test doesn't require Mendeley Desktop to run, only a blank document
	
    Dim fieldsList, bookmarksList

    ' Adds a field
    createField()

    bookmarksList = thisComponent.Bookmarks.ElementNames
    fieldsList = thisComponent.ReferenceMarks.ElementNames

    checkLengthLists(bookmarksList, -1, fieldsList, 0)

    ' Convert to bookmarks
    ZoteroUseBookmarks = True
    fnGetMarks(ZoteroUseBookmarks)

    bookmarksList = thisComponent.Bookmarks.ElementNames
    fieldsList = thisComponent.ReferenceMarks.ElementNames

    checkLengthLists(bookmarksList, 0, fieldsList, -1)
End Sub

Sub createField()
    Dim oSelection, oField

    oSelection = thisComponent.currentController.getViewCursor()

    oField = thisComponent.createInstance("com.sun.star.text.ReferenceMark")
    oField.setName ("ADDIN CSL_CITATION {this is some JSON} RND0123456789")
    oSelection.String = "(Smith 2009)"

    oSelection.text.insertTextContent(oSelection, oField, true)
End Sub

Sub checkLengthLists(list1, expectedLength1, list2, expectedLength2)
    If uBound(list1) <> expectedLength1 Then
        MsgBox("List1 unexpected uBound. Actual: " + uBound(list1) + " expected: " + expectedLength1)
        stop
    End If

    If uBound(list2) <> expectedLength2 Then
        MsgBox("List2 unexpected uBound. Actual: " + uBound(list2) + " expected: " + expectedLength2)
        stop
    End If
End Sub
