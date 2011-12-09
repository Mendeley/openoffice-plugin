
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009, 2010, 2011 Mendeley Ltd.
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
    outputPath = testsPath() & "output/"
End Function

Function crLf() As String
    crLf = Chr(13) & Chr(10)
End Function

Sub runUnitTests()
    unitTest = True
    
    Call testInsertCitation
    Call testEditCitation
    Call testMergeCitations
    Call testApplyFormatting
    Call testRefreshDocument
    
    ' TODO: port these tests from Word VBA to OpenOffice Basic
    'Call testChangeCitationStyle
    'Call testInsertBibliography
    'Call testExportOpenOfficeSimple
    'Call testExportWithoutFieldsSimple
    
    'Call Application.Quit
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

Function compareStrings(actual As String, expected As String, location As String) As Boolean
    Dim index As Long

    compareStrings = True

    'If Not (Len(actual) = Len(expected)) Then
    '    MsgBox "actual and expected lengths don't match: " & Len(actual) & ", " & Len(expected)
    '    compareStrings = False
    '    'Exit Function
    'End If

    Dim partialExpected As String
    Dim partialActual   As String
    Dim actualIndex As Long

    actualIndex = 0

    index = 0

ContinueLoop:    
    While index < Len(expected)
        index = index+1

        Dim actualChar As Integer
        Dim expectedChar As Integer
        
        partialExpected = mid(expected, index, 1)
        expectedChar = Asc(partialExpected)
        
        If expectedChar = 8288 Then
            Goto ContinueLoop
        End If
        
        ' skip Chr(8288) in input (using Replace() didn't work)
        actualChar = 8288
        While actualChar = 8288
            actualIndex = actualIndex + 1
            partialActual = mid(actual, actualIndex, 1)
            actualChar = Asc(partialActual)
        Wend
        
        If actualChar <> expectedChar Then
            MsgBox "Test failed at " & location & Chr(13) &_
                   "Character index: " & index & Chr(13) &_
                   "Actual Character: " & actualChar & Chr(13) &_
                   "Expected Character: " & expectedChar & Chr(13) &_
                   "Actual:   " & actual & Chr(13) &_
                   "Expected: " & expected
            compareStrings = False
            Exit Function
        End If
    Wend
    
    ' todo: check there's not more of the actual string left to parse, but ignoring the 8288 characters
End Function

' -- tests --

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
    Dim testsPath As String
    Dim documentName As String
    Dim outputPath As String
    
    Call initTests(testsPath, outputPath)

    Call newDocument()
    
    ' Add a citation and bibliogrphy
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    ActiveDocument.range.InsertAfter (" some more test" & vbCrLf & "New paragraph." & vbCrLf)
    
    Call ActiveDocument.Select
    Call Selection.Collapse(wdCollapseEnd)
    Call insertBibliography
    
    ' Export as expected text file
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportOO-expected.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
    
    Dim exportedFilename As String
    exportedFilename = outputPath & "exportOO-exported.doc"
    
    ' output as oo compatible
    Application.DisplayAlerts = wdAlertsNone
    Call privateExportCompatibleOpenOffice(exportedFilename)
    Application.DisplayAlerts = wdAlertsAll
    Call ActiveDocument.Close
    Application.Documents.Open (outputPath & "exportOO-exported.doc")
    
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportOO-actual.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
    
    Call ActiveDocument.Close
    
    ' clean up .doc so it's not read when doing the refreshDocuments test
    Call Kill(outputPath & "exportOO-exported.doc")
End Sub

Sub testExportWithoutFieldsSimple()
    Dim testsPath As String
    Dim documentName As String
    Dim outputPath As String
    
    Call initTests(testsPath, outputPath)
    
    oDesktop = createUnoService("com.sun.star.frame.Desktop")
    sUrl = "my_file.txt" 
    mFileProperties(0).Name = "FilterName" 
    mFileProperties(0).Value = "scalc: Text - txt - csv (StarCalc)" 
    mFileProperties(1).Name = "FilterFlags" 
    mFileProperties(1).Value = "FIX,,0,1,0/2/13/2/14/2/59/2/60/1" 
    oDocument =oDesktop.loadComponentFromURL(sUrl,"_blank",0,mFileProperties()) 
    
    ' Add a citation and bibliogrphy
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    ActiveDocument.range.InsertAfter (" some more test" & vbCrLf & "New paragraph." & vbCrLf)
    Call ActiveDocument.Select
    Call Selection.Collapse(wdCollapseEnd)
    Call insertBibliography
    
    ' Export as expected text file
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportWithoutFields-expected.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
    
    Dim exportedFilename As String
    exportedFilename = outputPath & "exportWithoutFields-exported.doc"
    
    ' output without mendeley fields
    Call removeMendeleyFields(ActiveDocument())
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(exportedFilename)
    Application.DisplayAlerts = wdAlertsAll
    Call ActiveDocument.Close
    Application.Documents.Open (outputPath & "exportWithoutFields-exported.doc")
    
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportWithoutFields-actual.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
    
    Call ActiveDocument.Close
    
    ' clean up .doc so it's not read when doing the refreshDocuments test
    Call Kill(outputPath & "exportWithoutFields-exported.doc")
End Sub

' this would be a better test since it's a real document with more complexity,
' but currently there's a bug which needs fixing (#17064) so it's disabled
Sub testExportOpenOfficeFromFile()
    Dim testsPath As String
    Dim documentName As String
    Dim outputPath As String
    
    Call initTests(testsPath, outputPath)
        
    ' search for all .docx files
    Dim filename
    filename = Dir(testsPath & "*.docx")
    
    If filename = "" Then
        MsgBox "Whoops, file not found"
    End If
    
    Set eventClassModuleInstance.App = Nothing
    Application.Documents.Open (testsPath & filename)
    
    ' refresh and export expected xml
    If Not refreshDocument(False) Then
        ActiveDocument.range.Text = "refreshDocument() failed"
    End If
    
    ' Export as expected text file
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportOO-expected.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
    
    Dim exportedFilename As String
    exportedFilename = outputPath & "exportOO-exported.doc"
    
    ' output as oo compatible
    Call privateExportCompatibleOpenOffice(exportedFilename)
    Call ActiveDocument.Close
    Application.Documents.Open (outputPath & "exportOO-exported.doc")
    
    Application.DisplayAlerts = wdAlertsNone
    Call ActiveDocument.SaveAs(outputPath & "exportOO-actual.txt", wdFormatText)
    Application.DisplayAlerts = wdAlertsAll
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
    Call compareStrings(documentText(), "expected", "changeCitation-APA")
    
    Call setCitationStyle("http://www.zotero.org/styles/ieee")
    Call refreshDocument(False)
    
    Call compareStrings(documentText(), "expected", "changeCitation-IEEE")
    
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
    Dim testsPath As String
    Dim outputPath As String
    Call initTests(testsPath, outputPath)
    
    ' uninitialize
    Set eventClassModuleInstance.App = Nothing

    Call newDocument()
    
    ' Add a citation at start
    Call privateInsertCitation("{80fd12bc-8c23-498c-a845-f29cd215dbec}")
    ActiveDocument.range.InsertAfter (" some more test" & vbCrLf & "New paragraph." & vbCrLf)
    Call ActiveDocument.Select
    Call Selection.Collapse(wdCollapseEnd)
    Call insertBibliography
    
    Call ActiveDocument.SaveAs(outputPath & "insertBibliography.txt", wdFormatText)
    Call ActiveDocument.Close
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
