

' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009, 2010, 2011 Mendeley Ltd.
' Copyright (c) 2006 Center for History and New Media
'                    George Mason University, Fairfax, Virginia, USA
'                    http://chnm.gmu.edu
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

' Thanks to the Zotero developers whose Word/Open Office plugin source code was
' frequently referred to and borrowed from in the development of this plugin

' author: steve.ridout@mendeley.com

Option Explicit


Function buildingPlugin() As Boolean
    buildingPlugin = Left(ActiveDocument.Name, Len(TEMPLATE_NAME_DURING_BUILD)) = TEMPLATE_NAME_DURING_BUILD
End Function

Function isMendeleyInstalled() As Boolean
    Dim myWS As Object
    Dim executablePath As String
    On Error GoTo ErrorHandler
    Set myWS = CreateObject("WScript.Shell")
    executablePath = myWS.RegRead("HKEY_CURRENT_USER\Software\Mendeley Ltd.\Mendeley Desktop\ExecutablePath")
    
    isMendeleyInstalled = (Dir(executablePath) <> "")
    Exit Function
  
ErrorHandler:
    isMendeleyInstalled = False
End Function

Sub reportError()
    Dim fixingInstructions As String

    Dim errorDescription
    Dim errorLine

    errorDescription = Error
    errorLine = Erl

    Dim vbCrLf
    vbCrLf = Chr(13)

    If isMendeleyInstalled() = False Then
        Exit Sub
    End If

	' todo: line number is specified but the source file is not - can we get it somehow?
    MsgBox "Error: " + errorDescription + Chr$(13) + "At line: " + errorLine
End Sub

Sub sendWordProcessorVersion()
      Call extGetStringResult(extSetWordProcessor("OpenOffice"))
End Sub

Function MakePropertyValue( Optional cName As String, Optional uValue ) As com.sun.star.beans.PropertyValue
   Dim oPropertyValue As New com.sun.star.beans.PropertyValue
   If Not IsMissing( cName ) Then
      oPropertyValue.Name = cName
   EndIf
   If Not IsMissing( uValue ) Then
      oPropertyValue.Value = uValue
   EndIf
   MakePropertyValue() = oPropertyValue
End Function

Function GetReadConfigAccess( ByVal nodePath As String ) As Object
    Dim configProvider as object
    configProvider = createUnoService( "com.sun.star.configuration.ConfigurationProvider" )
    Dim configAccess as object

    Dim args(0) as Variant
    Set args(0) = MakePropertyValue("nodepath",nodePath)

    configAccess = configProvider.createInstanceWithArguments( "com.sun.star.configuration.ConfigurationAccess", args )

    GetReadConfigAccess() = configAccess
End Function

Sub warnAboutAlwaysSaveAs
    ' If the document has already been saved and given a file extension,
    ' check that the extension is one that we support (e.g. ODT)
    If ThisComponent.getUrl <> "" Then
        If Not isThisFileOdf() Then
            Call showAlwaysSaveAsWarning
        End If
        Exit Sub
    End If

    dim configAccess as object
    configAccess = GetReadConfigAccess("/org.openoffice.Setup/Office/Factories/com.sun.star.text.TextDocument")
    dim defaultFileFormat as Variant
    defaultFileFormat = configAccess.getByName("ooSetupFactoryDefaultFilter")

    If defaultFileFormat = "writer8" Or defaultFileFormat = "writer8_template" Then
        ' this is fine
        Exit Sub
    End If

    Call showAlwaysSaveAsWarning
End Sub

Function isThisFileOdf As Boolean
    ' check if this file is in .odf format
    Dim fileExtension As String
    fileExtension = Right(ThisComponent.getUrl, 4)

    If Left(fileExtension, 1) = "." Then
      If (fileExtension = ".odt") Or (fileExtension = ".ott") Then
          ' Using safe format
          isThisFileOdf = True
          Exit Function
      End If
    End If
    isThisFileOdf = False
End Function

Sub showAlwaysSaveAsWarning
    Dim optionLocation As String
    Dim guiType As Long

    guiType = GetGUIType
    If guiType = 3 Or guiType = 4 Then
        ' Mac OSX
        optionLocation = "OpenOffice.org->Preferences...->Load/Save->General"
    Else
        ' Windows / Linux
        optionLocation = "Tools->Options...->Load/Save->General"
    End If

    MsgBox "IMPORTANT WARNING:" & Chr(13) & Chr(13) _
        & "Your Mendeley citations will not be saved with your document using the current file format. Please use File->Save As... " _
        & "to save your document as an ODF Text Document, or an " _
        & "ODF Text Document Template." & Chr(13) & Chr(13) _
        & "To avoid this message in future, you can change the ""Always Save As"" option in " _
        & optionLocation + " to ODF Text Document." & Chr(13) & Chr(13) _
        & "To save in .doc format please use the ""Export MS Word Compatible"" button in the Mendeley toolbar, " _
        & "instead of using File->Save or File->Save As.", 0, _
          "Mendeley OpenOffice Plugin"
End Sub

Function fileContainsLines(filePath As String, linesToFind() As String) As Boolean
    Dim nextLineToFind As Long
    nextLineToFind = LBound(linesToFind)

    Dim thisLine As String
    Dim fileId as Long
    fileId = FreeFile
    Open filePath For Input As #fileId 'Open the file for input
    Do While NOT EOF(fileId) 'While NOT End Of File
        Line Input #fileId, thisLine 'Read some data!

        ' ignore whitespace at start
        Do While Left(thisLine, 1) = " "
            thisLine = Right(thisLine, Len(thisLine) - 1)
        Loop

        If thisLine = linesToFind(nextLineToFind) Then
            nextLineToFind = nextLineToFind + 1
            If nextLineToFind > UBound(linesToFind) Then
                ' found all lines
                fileContainsLines = True
                GoTo EndOfFunction
            End If
        Else
            nextLineToFind = LBound(linesToFind)
        End If
    Loop

    fileContainsLines = False

EndOfFunction:
    Close #fileId
End Function

' Refresh the citations in this document and update the
' citation selector combo-box
'
' @param openingDocument Set to true if the refresh is being
' called whilst opening a new document or false if refreshing
' an existing already-open document
'
Function refreshDocument(Optional openingDocument As Boolean, Optional unitTest As Boolean)

    Dim currentDocumentPath As String
    currentDocumentPath = activeDocumentPath()

    refreshDocument = False
    
    ZoteroUseBookmarks = False
    
        Call warnAboutAlwaysSaveAs

    If openingDocument = True Then
        If isMendeleyRunning() = True Then
            Call setCitationStyle(getCitationStyleId())
        Else
            Exit Function
        End If
    End If
    
    If launchMendeleyIfNecessary() <> CONNECTION_CONNECTED Then
        Exit Function
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        Exit Function
    End If
    
	
    Call sendWordProcessorVersion
    
    ' update combo box
        ' Tell Mendeley Desktop what citation style to use
        Dim newStyle As String
        newStyle = getCitationStyleId()
        Call extGetStringResult(extSetCitationStyle(newStyle))
    
    ' Subscribe to events (e.g. WindowSelectionChange) doing on refreshDocument as it
    ' doesn't work in initialise() when addExternalFunctions() is also called

    Dim citationNumberCount As Long
    citationNumberCount = 0
    
    Dim bibliography As String

    Call extGetStringResult(extResetCitations)
    
    Dim marks
    marks = fnGetMarks(ZoteroUseBookmarks)
    
    Dim markName As String
    
        Dim thisField

    Dim mark

    Dim citationNumber As Long
    citationNumber = 0

    Dim oRange
    Dim oDupRange
    Dim empty_args(0) as new com.sun.star.beans.PropertyValue
    
    For Each mark In marks
        Set thisField = mark
        
        markName = getMarkName(thisField)
        
        If startsWith(markName, "ref Mendeley") Then
            markName = Right(markName, Len(markName) - 4)
            thisField.code.Text = markName
        End If
        
        If isMendeleyCitationField(markName) Then
            citationNumber = citationNumber + 1
            
            Call extGetStringResult(extAddCitation(addUnicodeTags(markName)))
            
            ' Just send an empty string if the displayed text is a temporary placeholder
            Dim displayedText As String
            displayedText = getMarkText(thisField)
            'displayedText = getMarkTextWithFormattingTags(thisField)
            If displayedText = INSERT_CITATION_TEXT Or displayedText = MERGING_TEXT Then
                displayedText = ""
            End If
            Call extGetStringResult(extAddFormattedCitation(addUnicodeTags(displayedText)))
            
        End If
    Next
    
    Dim oldCitationStyle As String
    oldCitationStyle = getCitationStyleId()
    
    ' Now that we've compiled the list of uuids, give them to Mendeley Desktop
    ' and tell it to format the citations and bibliography
    Dim status As String
    status = extGetStringResult(extFormatCitationsAndBibliography)
        
    If Left(status, Len("failed")) = "failed" Then
        Call extGetStringResult(extBringPluginToForeground)
        Exit Function
    End If
    
    citationNumber = 0
    
    marks = fnGetMarks(ZoteroUseBookmarks)
    For Each mark In marks 'ActiveDocument.Fields
    If currentDocumentPath <> activeDocumentPath() Then
        Exit Function
    End If

        Set thisField = mark
        Dim fieldText As String
        fieldText = ""
        markName = getMarkName(thisField)

        
        If (isMendeleyCitationField(markName)) Then
            Dim stringLen As Long
            stringLen = extGetFormattedCitation(citationNumber)
            fieldText = extGetStringResult(stringLen)

            Dim previousFormattedCitation As String
            stringLen = extGetPreviouslyFormattedCitation(citationNumber)
            previousFormattedCitation = formatUnicode(extGetStringResult(stringLen))
            
            Dim jsonData As String
            stringLen = extGetCitationJson(citationNumber)
            jsonData = extGetStringResult(stringLen)

            If currentDocumentPath <> activeDocumentPath() Then
                Exit Function
            End If
            Set thisField = fnRenameMark(thisField, jsonData)
            
            If fieldText <> addUnicodeTags(getMarkText(thisField)) Then
                If currentDocumentPath <> activeDocumentPath() Then
                    Exit Function
                End If
                
                ' if Mendeley sends us an empty field, leave it alone since we want to
                ' preserve the user's formatting options
                If fieldText <> "" Then
                    Call applyFormatting(fieldText, thisField)
                End If
            End If
            
            citationNumber = citationNumber + 1
        ElseIf isMendeleyBibliographyField(markName) Then
            If Not InStr(markName, CSL_BIBLIOGRAPHY) > 0 Then
                    mark = fnRenameMark(mark, markName & " " & CSL_BIBLIOGRAPHY)
            End If
        
            If bibliography = "" Then
                bibliography = bibliography + extGetStringResult(extGetFormattedBibliography())
            End If
                oRange = fnMarkRange(thisField)
                oRange.setPropertyValue("ParaFirstLineIndent",0)
                oRange.setPropertyValue("ParaLeftMargin",0)

                oDupRange = oRange.Text.createTextCursorByRange(oRange)

                dim currentFont as String
                dim currentHeight as Long

                currentFont = oRange.CharFontName
                currentHeight = oRange.CharHeight

                oDupRange.insertDocumentFromUrl(convertToUrl(bibliography),empty_args())

                oRange = fnMarkRange(thisField)
                oRange.CharFontName = currentFont
                oRange.CharHeight = currentHeight

                ' Remove the initial space and newline that Mendeley sends with the bibliography
                ' to workaround a problem in OpenOffice.org
                oDupRange = oRange.Text.createTextCursorByRange(oRange)

                oDupRange.goLeft(0,False)
                oDupRange.goRight(2,True)
                oDupRange.String = ""

            
            'fieldText = bibliography
            'Call applyFormatting(fieldText, thisField)
        End If
        
        If Not (fieldText = "") Then
            ' Put text in field
        End If
        
    Next
    
    If Not unitTest Then
        Dim newCitationStyle As String
        newCitationStyle = extGetStringResult(extGetCitationStyleId())
        
        If (newCitationStyle <> oldCitationStyle) Then
            ' set new citation style
            Call setCitationStyle(newCitationStyle)
            
        End If
        
          previouslySelectedField = getFieldAtSelection()
        If Not IsNull(previouslySelectedField) And Not IsEmpty(previouslySelectedField) Then
            previouslySelectedFieldResultText = getMarkText(previouslySelectedField)
        Else
            previouslySelectedFieldResultText = ""
        End If
    End If
    
	
    refreshDocument = True
End Function

Function appendJson(field As field, json As String) As field
    Dim markName As String
    
    markName = fnMarkName(field)
    
    ' remove Prev{} field
    Dim prevString As String
    prevString = getPreviousFormattedCitation(field)
    
    Dim position As Long
    position = InStr(markName, " Prev{")
    If position > 0 Then
        markName = Left(markName, position - 1)
    End If
    
    ' Add prev to json
    Dim jsonObject As Object
    Set jsonObject = jsonLibrary.parse(json)
    
    Dim mendeleyObject As Object
    Set mendeleyObject = new mendeleyDictionary
    Call mendeleyObject.insert("previousFormattedCitation", prevString)
    Call jsonObject.insert("mendeley", mendeleyObject)
    
    json = jsonLibrary.toString(jsonObject)
    
    ' append the document JSON to the mark name
    Set appendJson = fnRenameMark(field, markName & " CSL_CITATION " & json)
End Function

Function getPreviousFormattedCitation(mark)
    Dim markName As String
    Dim position As Long
    
    markName = getMarkName(mark)
    position = InStr(markName, " Prev{")
    
    If position > 0 Then
        getPreviousFormattedCitation = Mid(markName, position + 6, Len(markName) - position - 6)
    Else
        getPreviousFormattedCitation = ""
    End If
End Function

Sub setCitationStyle(style As String)
    Call subSetProperty(MENDELEY_CITATION_STYLE, style)
    Call extGetStringResult(extSetCitationStyle(style))
End Sub

Function getCitationStyleId() As String
    getCitationStyleId = fnGetProperty(MENDELEY_CITATION_STYLE)
    If getCitationStyleId = "" Then
        getCitationStyleId = DEFAULT_CITATION_STYLE
    End If
End Function

Function getStyleNameFromId(styleId As String) As String
    ' For compatibility with old system where the name was used as the identifier
    getStyleNameFromId = styleId
    
    If Not (startsWith(styleId, "http://") Or startsWith(styleId, "https://")) Then
        Exit Function
    End If
    
    Dim index As Long
    index = 0
    Dim currentStyleId As String
    
    currentStyleId = fnGetProperty(RECENT_STYLE_ID & " " & index)
    Do While currentStyleId <> ""
        If currentStyleId = styleId Then
            getStyleNameFromId = fnGetProperty(RECENT_STYLE_NAME & " " & index)
            Exit Function
        End If
        index = index + 1
        currentStyleId = fnGetProperty(RECENT_STYLE_ID & " " & index)
    Loop
End Function

' returns the user account which this document is currently linked to
Function mendeleyUserAccount() As String
    On Error GoTo CatchError
    
    mendeleyUserAccount = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value
    Exit Function
    
CatchError:
    mendeleyUserAccount = ""
End Function

Sub setMendeleyUserAccount(value As String)
    On Error GoTo CatchError
    
    Dim test As String
    
    ' if MENDELEY_DOCUMENT property not set this will throw an exception
    test = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value
    
    ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value = value

    Exit Sub
CatchError:
    ActiveDocument.CustomDocumentProperties.Add Name:=MENDELEY_USER_ACCOUNT, _
        LinkToContent:=False, value:=value, Type:=msoPropertyTypeString
End Sub


Function getFieldAtSelection()
      Dim oRange
      oRange = fnSelection()

      ' from Zotero src:
      Dim oBookmark, oRange1, oRange2, oField, oVC
      Dim oPortionEnum1, oPortion1, oPortionEnum2, oPortion2, sName As String

      oVC = thisComponent.currentController.viewCursor
      oRange1 = oRange.Text.createTextCursorByRange(oRange)
      oRange2 = oRange.Text.createTextCursorByRange(oRange)
      oRange1.goleft(1, False)
      oRange1.gotoStartOfParagraph (True)
      oRange2.gotoEndOfParagraph (True)
      oPortionEnum1 = oRange1.createEnumeration.nextElement.createEnumeration
      While oPortionEnum1.hasMoreElements
          oPortion1 = oPortionEnum1.nextElement
          If Not (fnOOoObject(oPortion1) Is Nothing) And oPortion1.isStart Then
              sName = fnOOoObject(oPortion1).Name
              If (InStr(sName, MENDELEY_CITATION) > 0) Or (InStr(sName, MENDELEY_BIBLIOGRAPHY) > 0) Or (InStr(sName, MENDELEY_EDITED_CITATION) > 0) _
                      Or (InStr(sName, CSL_CITATION) > 0) Or (InStr(sName, CSL_BIBLIOGRAPHY) > 0) Then
                  oPortionEnum2 = oRange2.createEnumeration.nextElement.createEnumeration
                  While oPortionEnum2.hasMoreElements
                      oPortion2 = oPortionEnum2.nextElement
                      If Not (fnOOoObject(oPortion2) Is Nothing) And Not oPortion2.isStart Then
                          If fnOOoObject(oPortion2).Name = sName Then
                              Set getFieldAtSelection = fnOOoObject(oPortion2)
                              'If bSelect Then subSelect(getFieldAtSelection)
                          End If
                      End If
                  Wend
              End If
          End If
      Wend
End Function

' Returns connection status
Function launchMendeleyIfNecessary() As Long
    ' Only need to launch Mendeley if the document contains the
    ' MENDELEY property
    If Not isMendeleyDocument Then
        launchMendeleyIfNecessary = CONNECTION_NOT_A_MENDELEY_DOCUMENT
        Exit Function
    End If

    launchMendeleyIfNecessary = extCheckConnectionAndSetupIfNeeded()
    
    If launchMendeleyIfNecessary = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND Then
        launchMendeleyIfNecessary = extLaunchMendeley()
    End If
End Function

Function isMendeleyRunning() As Boolean
    ' In OpenOffice we cannot know if it's running in that point, so tries to connect
    ' and if possible to connect it's running
    isMendeleyRunning = True
End Function

Function isMendeleyDocument() As Boolean
    If fnGetProperty(MENDELEY_DOCUMENT) = "True" Then
        isMendeleyDocument = True
    Else
        isMendeleyDocument = False
    End If
End Function

Sub setMendeleyDocument(value As Boolean)
    If value Then
        Call subSetProperty(MENDELEY_DOCUMENT, "True")
    Else
        Call subSetProperty(MENDELEY_DOCUMENT, "False")
    End If
End Sub


Function isDocumentLinkedToCurrentUser() As Boolean
    Dim currentMendeleyUser As String
    Dim thisDocumentUser As String
    
    currentMendeleyUser = extGetStringResult(extGetUserAccount())
    thisDocumentUser = fnGetProperty(MENDELEY_USER_ACCOUNT)
    
    ' remove server protocol from account string
    thisDocumentUser = Replace(thisDocumentUser, "http://", "")
    thisDocumentUser = Replace(thisDocumentUser, "https://", "")
    
    If currentMendeleyUser = thisDocumentUser Then
        isDocumentLinkedToCurrentUser = True
    Else
        Dim result ' As VbMsgBoxResult
        
        Dim vbCrLf
        vbCrLf = Chr(13)
        
        If thisDocumentUser = "" Then
            ' if no user currently linked then set without asking user
            result = MSGBOX_RESULT_YES
        Else
            ' ask user if they want to link the document to their account
            result = MsgBox("This document has been edited by another Mendeley user: " + thisDocumentUser + vbCrLf + vbCrLf + _
                "Do you wish to enable the Mendeley plugin to edit the citations and bibliography yourself?" + vbCrLf + vbCrLf, _
                MSGBOX_BUTTONS_YES_NO, "Enable Mendeley plugin for this document?")
        End If

        If result = MSGBOX_RESULT_YES Then
            Call subSetProperty(MENDELEY_USER_ACCOUNT, currentMendeleyUser)
            isDocumentLinkedToCurrentUser = True
        Else
            isDocumentLinkedToCurrentUser = False
        End If
    End If
End Function


' Returns true if mainString starts with the subString
Function startsWith(mainString As String, subString As String) As Boolean
    startsWith = Left(mainString, Len(subString)) = subString
End Function

Sub applyFormatting(markup As String, mark)
    ' parse range and apply following formatting:
    ' <i></i> italics
    ' <b></b> bold
    ' <u></u> underline
    ' <sup></sup> superscript
    ' <sub></sub> subscript

    ' add extra space at start because the Range.Delete function will
    ' delete the whole field if we attempt to delete the first character
    ' (it gets deleted later)

        mark = subSetMarkText(mark, markup)
    
        Dim range
        Set range = fnMarkRange(mark)

        Dim subRange 'As range
        Set subRange = range.Text.createTextCursorByRange(range)
        range.setPropertyValue("CharEscapement", 0)
        range.setPropertyValue("CharEscapementHeight", 100)
    
    Dim startPosition As Long
    Dim endPosition As Long
    
        startPosition = 0
        endPosition = Len(range.Text.String)
    
    Call applyStyleToTagPairs("i", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("b", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("u", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("sup", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("sub", subRange, startPosition, endPosition)
    
    If InStr(getMarkText(mark), "second-field-align") > 0 Or InStr(getMarkText(mark), "hanging-indent") Then
            range.setPropertyValue("ParaFirstLineIndent", 0)
            range.setPropertyValue("ParaLeftMargin", 0)
    End If
    
    Call applyStyleToTagPairs("second-field-align", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("hanging-indent", subRange, startPosition, endPosition)
    
    ' insert unicode characters
    ' (would be better if we could send unicode characters directly over the local socket,
    '  but it doesn't seem to work)
    Call applyStyleToTagPairs("unicode", subRange, startPosition, endPosition)

    ' Add paragraph breaks in place of <p>
    Call applyStyleToIndividualTags("p", subRange, startPosition, endPosition)
End Sub

Function formatUnicode(inputText As String) As String
    ' Convert all <unicode>???</unicode> tags to unicode characters
    formatUnicode = ""
    If inputText = "" Then
        Exit Function
    End If
    
    Dim result As String
    result = ""
    
    Dim positionInInputText As Long
    positionInInputText = 1

    Dim startPosition As Long
    Dim endPosition As Long
    
    startPosition = InStr(inputText, "<unicode>")
    endPosition = 1
    
    Do While startPosition > 0
        result = result & Mid(inputText, positionInInputText, startPosition - positionInInputText)
        
        startPosition = startPosition + Len("<unicode>")
        endPosition = InStr(startPosition, inputText, "</unicode>")
        
        Dim charCode As String
        charCode = Mid(inputText, startPosition, endPosition - startPosition)
        
            result = result & Chr(charCode)
        
        endPosition = endPosition + Len("</unicode>")
        positionInInputText = endPosition
        startPosition = InStr(endPosition, inputText, "<unicode>")
    Loop
    
    Dim length As Long
    length = Len(inputText)
        
    ' add the rest
    formatUnicode = result & Mid(inputText, endPosition, length - endPosition + 1)
End Function

Function addUnicodeTags(inputString As String) As String
        Dim outputString
    Dim position As Long
    Dim charCode As Long
    Dim stringToAppend As String
    
    ' use DynamicStringType to avoid inefficient string concatenations
    outputString = DynamicStringInit(Len(inputString))

    Dim appendFrom As Long
    appendFrom = 1
    
    ' use DynamicStringType to avoid inefficient string concatenations
    outputString = DynamicStringInit(Len(inputString))
    
    Dim outputStringPosition As Long
    outputStringPosition = 1
    
    For position = 1 To Len(inputString)
            charCode = Asc(Mid(inputString, position, 1))
        
        If charCode < 0 Then
            charCode = 65536 + charCode
        End If
        
        If charCode >= 128 Then
            outputString = DynamicStringAppend(outputString, Mid$(inputString, appendFrom, position - appendFrom))
            outputString = DynamicStringAppend(outputString, "<unicode>" & charCode & "</unicode>")
            appendFrom = position + 1
        End If
    Next
    
    If appendFrom < position Then
        outputString = DynamicStringAppend(outputString, Mid$(inputString, appendFrom, position - appendFrom))
    End If
    
    ' remove the extra allocated space from the end of outputString
    addUnicodeTags = DynamicStringGet(outputString)
End Function


Sub testAddUnicodeTags()
    Dim testString As String
    
    testString = "Hello " & Chr(181)
    
    MsgBox "before: " & testString
    testString = addUnicodeTags(testString)
    
    MsgBox "after: " & testString
End Sub

Sub applyStyleToTagPairs(tag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long)

    Dim startTag As String
    Dim endTag As String
    
        Dim thisRange 'As Range
        Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
    
    startTag = "<" + tag + ">"
    endTag = "</" + tag + ">"
    
    ' Maximum number of characters used in the first field
    ' Used for setting the second-field-align tab stopa
    Dim maxFirstFieldLength As Long
    maxFirstFieldLength = 0

    Do While Not (rangeString(thisRange) = "") And Not (InStr(rangeString(thisRange), startTag) = 0)
        ' find and remove start tag
            Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
        
        Dim startTagPosition As Long

            startTagPosition = InStr(fnReplace(rangeString(thisRange), Chr(13), ""), startTag) - 1
            thisRange.goRight(startTagPosition, False)
            thisRange.goRight(2 + Len(tag), True)
            thisRange.String = ""
        
        ' find and remove end tag
        
          Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)

        Dim endTagPosition As Long
        
            endTagPosition = InStr(fnReplace(rangeString(thisRange), Chr(13), ""), endTag) - 1
            thisRange.goRight(endTagPosition, False)
            thisRange.goRight(3 + Len(tag), True)
            ' add a space on the end, which is deleted later. Otherwise, if we are
            ' at the end of a mark, new characters will be inserted OUTSIDE the mark
            thisRange.String = " "

            Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
            thisRange.goRight(startTagPosition, False)
            thisRange.goRight(endTagPosition - startTagPosition, True)
        
        ' apply style
        Select Case tag
                Case "b"
                    thisRange.setPropertyValue("CharWeight", 150)
                Case "i"
                    thisRange.setPropertyValue("CharPosture", 2)
                Case "u"
                    thisRange.setPropertyValue("CharUnderline", 1)
                Case "sup"
                    thisRange.setPropertyValue("CharEscapement", 33)
                    thisRange.setPropertyValue("CharEscapementHeight", 58)
                Case "sub"
                    thisRange.setPropertyValue("CharEscapement", -33)
                    thisRange.setPropertyValue("CharEscapementHeight", 58)
                Case "second-field-align"
                    ' Remove spaces at the end of the range
                    ' (@todo remove this if fixed in Mendeley Desktop)
                    Do While Right(rangeString(thisRange), 1) = " "
                      Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
                      thisRange.collapseToStart
                      thisRange.goRight(endTagPosition-1,False)
                      thisRange.goRight(1, True)
                      thisRange.String = ""
                      endTagPosition = endTagPosition - 1
                    Loop

                    If ((endTagPosition - startTagPosition) > maxFirstFieldLength) Then
                        maxFirstFieldLength = endTagPosition - startTagPosition
                    End If

                    ' remove subsequent spaces after the range
                    Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
                    Call thisRange.collapseToStart
                    thisRange.goRight(endTagPosition,False)
                    thisRange.goRight(1, True)

                    Do While rangeString(thisRange) = " "
                        thisRange.String = ""
                        Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
                        thisRange.collapseToStart
                        thisRange.goRight(endTagPosition,False)
                        thisRange.goRight(1,True)
                    Loop
                    ' Insert tab
                    Call thisRange.collapseToStart
                    thisRange.String = Chr(9)
                Case "hanging-indent"
                    Call setHangingIndent(wholeRange)

                Case "unicode"
                    Dim characterCode As Long
                    characterCode = thisRange.String
                    thisRange.String = Chr(characterCode)
        End Select

            ' remove extra space from after the tag
            Call thisRange.collapseToEnd()
            thisRange.goRight(1, True)
            thisRange.String = ""

            Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
    Loop
    
    If tag = "second-field-align" And maxFirstFieldLength > 0 Then
        ' Create tab stop - the position is calculated approximately
        ' and works reasonably well for numbered citations
        Call setHangingIndent(wholeRange, 4 + 6 * maxFirstFieldLength)
    End If
    
End Sub

Sub setHangingIndent(range As range, Optional length As Long)
        If IsMissing(length) Then
            length = 40
        End If
        range.setPropertyValue("ParaLeftMargin", Int(TWIPS_TO_100TH_MM * length * 20))
        range.setPropertyValue("ParaFirstLineIndent", Int(-TWIPS_TO_100TH_MM * length * 20))
End Sub

Function rangeString(range As range) As String
        rangeString = range.String
End Function

Sub applyStyleToIndividualTags(tag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long)

    Dim startTag As String
    Dim endTag As String
    
    startTag = "<" + tag + ">"
    endTag = "</" + tag + ">"
    
        Dim thisRange
        Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)

    Do While Not (rangeString(thisRange) = "") And Not (InStr(rangeString(thisRange), startTag) = 0)
        ' find and remove start tag
        Dim startTagPosition As Long
        
            startTagPosition = InStr(fnReplace(rangeString(thisRange), Chr(13), ""), startTag) - 1
            thisRange.goRight(startTagPosition, False)
            thisRange.goRight(2 + Len(tag), True)
            thisRange.String = ""

        ' apply formatting
        Select Case tag
            Case "p"
              thisRange.String = Chr(13)
        End Select
        
            Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
    Loop
End Sub

Sub checkForCitationEdit()
    ' We can't update the document until we've updated the previously
    ' selected field, so if the user cancels an edit, needUpdate is set
    ' to true and we UpdateDocument() at the end
    Dim needUpdate As Boolean
    needUpdate = False
    
    ' TODO: deal with edits of bibliographies
    Dim objectExists As Boolean
  objectExists = True
    If Not (previouslySelectedField Is Nothing) And Not IsMissing(previouslySelectedField) And Not IsEmpty(previouslySelectedField) And objectExists Then
        If Not (previouslySelectedField.result.Text = previouslySelectedFieldResultText) Then
            
            If Not isMendeleyRunning() Then
                ' Don't need do anything - this edit will get detected next time the user refreshes
                Exit Sub
            End If
            
            Dim markName As String
            markName = getMarkName(previouslySelectedField)
            
            Dim displayedText As String
            displayedText = addUnicodeTags(getMarkText(previouslySelectedField))
            displayedText = Replace(displayedText, Chr(13), "<p>")
            Call extGetStringResult(extSetDisplayedText(displayedText))
            
            Dim newMarkName As String
            newMarkName = extGetStringResult(extCheckManualFormatAndGetFieldCode(markName))
            
            If markName <> newMarkName Then
                Call fnRenameMark(previouslySelectedField, newMarkName)
                ' Disabled until we can send rich text formatting tags from the displayed
                ' citations to Mendeley Desktop
                'displayedText = extGetStringResult(extGetDisplayedText())
                'Call subSetMarkText(previouslySelectedField, displayedText)
                needUpdate = True
            End If
                        
        End If
    End If
    
    If needUpdate Then
        Call refreshDocument
    End If
End Sub

' ---------------------------
'    Utility Functions
' ---------------------------

' Returns the index of item in the array, or -1 if not found
' (doesn't permit arrays with -ve lower bound)
Function indexOf(container() As String, item As String) As Long
    Dim index As Long
    
    If LBound(container) < 0 Then
        MsgBox "indexOf doesn't permit lower bounds < 0"
        Exit Function
    End If
    
    For index = LBound(container) To UBound(container)
        If container(index) = item Then
            indexOf = index
            Exit Function
        End If
    Next
    
    ' not found
    indexOf = -1
End Function

' ----- Functions from Zotero -----
' All the following functions were originally copied from the Zotero code base
' (https://www.zotero.org/svn/integration/ice/trunk/plugin.vb - revision 3444)
' and have been modified over the period between September 2008 to July 2009
' in order to work properly with the rest of the Mendeley code

' Gets a tag from the document properties
Function fnGetProperty(sProperty As String) As String
    Dim sPropertyName As String, i As Long, sPropertyValue As String
    i = 0
    
    On Error GoTo EndOfFunction
    
    While (True)
        i = i + 1
        sPropertyName = sProperty & "_" & i
        
            sPropertyValue = thisComponent.DocumentProperties.getUserDefinedProperties().getPropertyValue(sPropertyName)
            If sPropertyValue = "" Then GoTo EndOfFunction
            fnGetProperty = fnGetProperty & sPropertyValue
    Wend
EndOfFunction:
    Exit Function
End Function

' Sets a tag from the document properties
Function subSetProperty(sProperty As String, ByVal sValue As String) As String
    Dim nProperties As Long, nIterations As Long, i As Long
    Dim sPropertyName As String, sPropertyValue As String
    
    
    If sValue = "" Then
        nProperties = 0
    Else
        ' fake ceil function
        nProperties = -Int(-Len(sValue) / MAX_PROPERTY_LENGTH)
        
        For i = 1 To nProperties
            sPropertyName = sProperty & "_" & i
            sPropertyValue = Mid(sValue, (i - 1) * MAX_PROPERTY_LENGTH + 1, MAX_PROPERTY_LENGTH)
            
                On Error Resume Next
                thisComponent.DocumentProperties.getUserDefinedProperties().addProperty(sPropertyName, 0, "")
                On Error GoTo 0
                thisComponent.DocumentProperties.getUserDefinedProperties().setPropertyValue(sPropertyName, sPropertyValue)
        Next
    End If
    
    ' delete extra fields
    On Error GoTo EndOfFunction
    While (True)
        nProperties = nProperties + 1
        sPropertyName = sProperty & "_" & nProperties
        
            thisComponent.DocumentProperties.getUserDefinedProperties().removeProperty (sPropertyName)
    Wend
    Exit Function
AddProperty:
    ActiveDocument.CustomDocumentProperties.Add Name:=sPropertyName, _
        LinkToContent:=False, value:=sPropertyValue, Type:=msoPropertyTypeString
    Resume Next
EndOfFunction:
End Function

'Returns the current selection
Function fnSelection()
        Dim oVC
        Set oVC = thisComponent.currentController.viewCursor
        fnSelection = oVC.Text.createTextCursorByRange(oVC)
End Function

Function fnOOoObject(oObject) As Object
    If Not IsNull(oObject.bookmark) And Not IsEmpty(oObject.bookmark) Then
        fnOOoObject = oObject.bookmark
    ElseIf Not IsNull(oObject.ReferenceMark) And Not IsEmpty(oObject.ReferenceMark) Then
        fnOOoObject = oObject.ReferenceMark
    ElseIf Not IsNull(oObject.TextSection) And Not IsEmpty(oObject.TextSection) Then
        fnOOoObject = oObject.TextSection
    End If
End Function

' Adds a mark (field, bookmark, or reference mark) with name sName at location oRange
Function fnAddMark(oRange, sName As String) 'As Field
    Dim i As Long, sBasename As String, oField
    Dim nOldStoryRange As Long, nOldField As Long
    Dim nNewStoryRange As Long, nNewField As Long
    Dim oBookmarks, sMarkText As String
    
    ' Escape if necessary
    If ZoteroUseBookmarks Then
        sName = fnSetFullBookmarkName(sName)
    End If
    
    ' Add Zotero mark
    If InStr(sName, MENDELEY_BIBLIOGRAPHY) > 0 Then
        sMarkText = BIBLIOGRAPHY_TEXT
    ElseIf InStr(sName, MENDELEY_CITATION_EDITOR) > 0 Then
        sMarkText = CITATION_EDIT_TEXT
        sName = MENDELEY_CITATION
    Else
        sMarkText = INSERT_CITATION_TEXT
    End If
    
        If ZoteroUseBookmarks Then
            'OOo automatically creates a unique bookmark name allowing for more than one bibliography
            oField = thisComponent.createInstance("com.sun.star.text.Bookmark")
        ElseIf isMendeleyBibliographyField(sName) Then
            'Create a document section for the bibliography
            sName = sName & REFERENCEMARK_RANDOM_DATA_SEPARATOR & fnGenerateRandomString(REFERENCEMARK_RANDOM_STRING_LENGTH)
            oField = thisComponent.createInstance("com.sun.star.text.TextSection")
        Else
            'Allow for more than one citation of same item in same position by creating a unique ReferenceMark name
            sName = sName & REFERENCEMARK_RANDOM_DATA_SEPARATOR & fnGenerateRandomString(REFERENCEMARK_RANDOM_STRING_LENGTH)
            oField = thisComponent.createInstance("com.sun.star.text.ReferenceMark")
        End If
        oField.setName (sName)
        oRange.String = sMarkText
        oRange.text.insertTextContent(oRange, oField, true)
        Set fnAddMark = oField
End Function

' Saves sName in document properties, if necessary, then returns the appropriate bookmark name
Function fnSetFullBookmarkName(sName) As String
    Dim nStringLength As Long
    
    ' Set property
    fnSetFullBookmarkName = ZOTERO_BOOKMARK_REFERENCE_PROPERTY & "_" & fnGenerateRandomString(BOOKMARK_ID_STRING_LENGTH)
    Call subSetProperty(fnSetFullBookmarkName, sName)
End Function

Function fnRenameMark(oMark, sNewname As String)
        If oMark.supportsService("com.sun.star.text.Bookmark") Then
            If (Left(oMark.Name, 11) = "ZOTERO_BREF") Then
                Call subSetProperty(oMark.Name, sNewname)
            Else
                oMark.Name = fnSetFullBookmarkName(sNewname)
            End If
        ElseIf oMark.supportsService("com.sun.star.text.TextSection") Then
            oMark.Name = sNewname
        ElseIf oMark.supportsService("com.sun.star.text.ReferenceMark") Then
            'The only way of renaming a referencemark is to delete it and recreate it with the new name
            'The process below loses any text formatting.
            Dim sString As String, oRange, oField

            oRange = oMark.Anchor


            ' Next two lines are on purpose: when the field is in a
      ' table if we don't force a read from oRange.Text.String
      ' the line insertTextContent fails because oRange.text is NULL
      ' (!!!!) (at least in OO.org 3.1.1 Build 9420 in Linux)
            Dim ForceRead as String
            ForceRead = oRange.Text.String

            sString = oRange.String
            oRange.String = sString
            oField = thisComponent.createInstance("com.sun.star.text.ReferenceMark")
            oField.Name = sNewname & REFERENCEMARK_RANDOM_DATA_SEPARATOR & fnGenerateRandomString(REFERENCEMARK_RANDOM_STRING_LENGTH)
            oRange.text.insertTextContent(oRange, oField, true)
            Set oMark = oField
        End If
    Set fnRenameMark = oMark
End Function

Function getMarkName(mark) As String
        If mark.supportsService("com.sun.star.text.ReferenceMark") Then
            Dim length As Long
            Dim separator As String
            separator = REFERENCEMARK_RANDOM_DATA_SEPARATOR
            getMarkName = mark.Name
            length = Len(getMarkName) - Len(separator) - REFERENCEMARK_RANDOM_STRING_LENGTH
            getMarkName = Left(getMarkName, length)
        Else
            getMarkName = mark.Name
        End If
End Function

Function getMarkText(mark) As String
        getMarkText = mark.Anchor.String
End Function

' NOTE: getMarkTextWithFormattingTags() and getTaggedRichTextFromRange() are not being
'       currently used because they are too slow to use on all citations in a document.
'       Keeping the code here just in case they can be optimised, or for occassional use
Function getMarkTextWithFormattingTags(mark) As String
        ' Not implemented for OpenOffice!!
        getMarkTextWithFormattingTags = mark.Anchor.String
End Function


Sub randomizeIfNeccessary()
    If seedGenerated Then
        Exit Sub
    End If

    ' generates a seed between -32500 and 32500
    Dim longNumber As Long
    longNumber = 24& * 60 * 60 * Day(Date) + 60& * 60 * Hour(Time()) + 60& * Minute(Time()) + Second(Time())
    
    Dim integerSeed As Integer
    integerSeed = (longNumber Mod 65000) - 32500
    
    Call Randomize(integerSeed)
    
    seedGenerated = True
End Sub

' Generates a random string
Function fnGenerateRandomString(nLength) As String
    Dim i As Long, nRandom As Long
    
    Call randomizeIfNeccessary
    
    For i = 1 To nLength
        ' use alphanumerics
        nRandom = Int((26 * 2 + 10) * Rnd)
        If (nRandom < 10) Then
            nRandom = nRandom + 48
        ElseIf (nRandom < (26 + 10)) Then
            nRandom = nRandom + 65 - 10
        Else
            nRandom = nRandom + 97 - (26 + 10)
        End If
        fnGenerateRandomString = fnGenerateRandomString & Chr(nRandom)
    Next
End Function


Function fnGetMarks(bBookmarks As Boolean)
    Dim i As Long, j As Long, mMarks()
    Dim oRef, oStoryRange, nCount As Long
    Dim textFrameStoryRange As Variant

    i = 0
    nCount = 0
        nCount = thisComponent.Bookmarks.Count + thisComponent.ReferenceMarks.Count + thisComponent.TextSections.Count
        Dim mBookmarks, mReferenceMarks, mTextSections
        ReDim mMarks(nCount)
        mBookmarks = thisComponent.Bookmarks.ElementNames
        mReferenceMarks = thisComponent.ReferenceMarks.ElementNames
        mTextSections = thisComponent.TextSections.ElementNames
        For j = 0 To UBound(mBookmarks)
            oRef = thisComponent.Bookmarks.getByname(mBookmarks(j))
            If (Left(mBookmarks(j), 9) = "Mendeley_" Or InStr(mBookmarks(j), "CSL_CITATION") > 0 Or InStr(mBookmarks(j), "CSL_BIBLIOGRAPHY") > 0) Then
                If bBookmarks Then
                    Set mMarks(i) = oRef
                Else
                    Set mMarks(i) = fnConvert(oRef)
                End If
                i = i + 1
            End If
        Next
        For j = 0 To UBound(mReferenceMarks)
            oRef = thisComponent.ReferenceMarks.getByname(mReferenceMarks(j))
            If (Left(mReferenceMarks(j), 9) = "Mendeley " Or InStr(mReferenceMarks(j), "CSL_CITATION") > 0 Or InStr(mReferenceMarks(j), "CSL_BIBLIOGRAPHY") > 0)  Then
                If Not bBookmarks Then
                    Set mMarks(i) = oRef
                Else
                    Set mMarks(i) = fnConvert(oRef)
                End If
                i = i + 1
            End If
        Next
        For j = 0 To UBound(mTextSections)
            oRef = thisComponent.TextSections.getByname(mTextSections(j))
            If (Left(mTextSections(j), 9) = "Mendeley " Or InStr(mTextSections(j), "CSL_CITATION") > 0 Or InStr(mTextSections(j), "CSL_BIBLIOGRAPHY") > 0) Then
                If Not bBookmarks Then
                    Set mMarks(i) = oRef
                Else
                    Set mMarks(i) = fnConvert(oRef)
                End If
                i = i + 1
            End If
        Next
    If i = 0 Then
        fnGetMarks = Array()
    Else
        ReDim Preserve mMarks(i - 1)
          Call subShellSort(mMarks)
        fnGetMarks = mMarks()
    End If
End Function

Sub subShellSort(mArray)
    'Based on a routine available from: http://www.oopweb.com/Algorithms/Documents/Sman/Volume/s_vss.txt
    Dim n As Long, h As Long, i As Long, j As Long, t, Ub As Long, LB As Long
    Dim nSR As Long, nFN As Long
    
    LB = LBound(mArray)
    Ub = UBound(mArray)
    
    ' compute largest increment
    n = Ub - LB + 1
    h = 1
    If n > 14 Then
        Do While h < n
            h = 3 * h + 1
        Loop
        h = h \ 3
        h = h \ 3
    End If
    Do While h > 0
    ' sort by insertion in increments of h
        For i = LB + h To Ub
            t = mArray(i)
            For j = i - h To LB Step -h
                If fnRangeComp(mArray(j), t) Then Exit For
                mArray(j + h) = mArray(j)
            Next j
            mArray(j + h) = t
        Next i
        h = h \ 3
    Loop
End Sub

Function fnRangeComp(m1, m2) As Boolean
    'Needed for the sort routine in OOo to get marks in correct order.
    Dim oR1, oR2, nLT1 As Long, nLT2 As Long
    
    Dim currentSelection
    Dim previousSelection
    Dim range1TableSelection
    Dim range2TableSelection
    
    Dim beforeRange1
    Dim beforeRange2
    
    oR1 = fnMarkRange(m1)
    oR2 = fnMarkRange(m2)
    nLT1 = fnLocationType(oR1)
    nLT2 = fnLocationType(oR2)
    
    currentSelection = ThisComponent.getCurrentController().getViewCursor()
    ' remember the current selection in-case we have to move it to find a table position
    previousSelection = currentSelection.getText().createTextCursorByRange(currentSelection)
    
    If nLT1 = ZOTERO_TABLE Then
        ThisComponent.getCurrentController().Select (oR1.TextTable)
        currentSelection.goLeft(1,False)
        range1TableSelection = currentSelection.getText().createTextCursorByRange(currentSelection)
    End If
    
    If nLT2 = ZOTERO_TABLE Then
        ThisComponent.getCurrentController().Select (oR2.TextTable)
        currentSelection.goLeft(1,False)
        range2TableSelection = currentSelection.getText().createTextCursorByRange(currentSelection)
    End If
    
    If nLT1 = ZOTERO_TABLE And nLT2 = ZOTERO_TABLE Then
        If oR1.TextTable.LinkDisplayName = oR2.TextTable.LinkDisplayName Then
            If IsEmpty(oR1.Cell) = False And IsEmpty(oR2.Cell) = False Then
                Dim cellNameCompare As Long
                cellNameCompare = isCellNameBigger(oR1.Cell.CellName, oR2.Cell.CellName)
                
                ' If both in same cell need to compare
                If cellNameCompare = True Or cellNameCompare = False Then
                    fnRangeComp = cellNameCompare
                Else
                    fnRangeComp = oR1.Cell.Text.compareRegionStarts(oR1, oR2)
                End If
            End If
        Else
            Dim temp
        
            ' Just creating a valid range to use as an argument for calling compareRegionStarts
            ' Not important that it's the document start
            Dim docStart
            docStart = ThisComponent.Text.getStart()
            
            ' if compareRegionStarts() throws an exception we know that the table is at the start of the document
            On Error GoTo Range1AtStart
            temp = ThisComponent.Text.compareRegionStarts(docStart, range1TableSelection)
            On Error GoTo Range2AtStart
            temp = ThisComponent.Text.compareRegionStarts(docStart, range2TableSelection)
        
            fnRangeComp = ThisComponent.Text.compareRegionStarts(range1TableSelection, range2TableSelection) > 0
        End If
    ElseIf nLT1 = ZOTERO_TABLE And nLT1 <> nLT2 Then
        On Error GoTo Range1AtStart
        fnRangeComp = ThisComponent.Text.compareRegionStarts(range1TableSelection, oR2) > 0
    ElseIf nLT2 = ZOTERO_TABLE And nLT1 <> nLT2 Then
        On Error GoTo Range2AtStart
        fnRangeComp = ThisComponent.Text.compareRegionStarts(oR1, range2TableSelection) > 0
    ElseIf nLT1 > nLT2 Then
        fnRangeComp = True
    ElseIf nLT1 < nLT2 Then
        fnRangeComp = False
    ElseIf nLT1 = ZOTERO_MAIN Then
        fnRangeComp = ThisComponent.Text.compareRegionStarts(oR1, oR2) > 0
    Else
        fnRangeComp = True
    End If
    
    ThisComponent.getCurrentController().Select (previousSelection)
    Exit Function
Range1AtStart:
        fnRangeComp = True
        ThisComponent.getCurrentController().Select (previousSelection)
    Exit Function
Range2AtStart:
        fnRangeComp = False
        ThisComponent.getCurrentController().Select (previousSelection)
End Function

' Receives things like A3, B6 and returns True if A3 is Bigger than B6
' (so if it appears later in the document)
Function isCellNameBigger(Cell1Name As String, Cell2Name As String) As Boolean
    ' Right now only works with A-Z columns and 0-9 rows
    isCellNameBigger = Cell1Name < Cell2Name
End Function

Function fnMarkRange(oMark)
    Set fnMarkRange = oMark.Anchor
End Function

'Depending on the location of the insertion point returns:
'  ZOTERO_MAIN (1) if in the main body of the document
'  ZOTERO_FOOTNOTE (2) if in a Footnote
'  ZOTERO_ENDNOTE (3) if in an Endnote
'  ZOTERO_ERROR otherwise
Function fnLocationType(oRange)
    Dim nLocation As Long, oText
        On Error GoTo InvalidSelection
        oText = oRange.Text
        On Error GoTo 0
        If oText.getImplementationName = "SwXCell" Then
            ' oText = oText.createTextCursor.textTable.Anchor.Text
            nLocation = ZOTERO_TABLE
        Else
            Select Case oText.getImplementationName
            Case "SwXBodyText"
                nLocation = ZOTERO_MAIN
            Case "SwXFootnote"
                If oText.supportsService("com.sun.star.text.Endnote") Then
                    nLocation = ZOTERO_ENDNOTE
                Else
                    nLocation = ZOTERO_FOOTNOTE
                End If
            Case Else
                nLocation = ZOTERO_ERROR
            End Select
        End If
    fnLocationType = nLocation
    Exit Function
InvalidSelection:
    fnLocationType = ZOTERO_ERROR
End Function

Sub subSetRangeText(oRange, sReplace As String)
        oRange.String = sReplace
End Sub

' This function, in OpenOffice, doesn't guarantee that the modified oMark is the same that it receives,
' the callers should do oMark = subSetMarkText(oMark, "test")
' And yes, right, pass by reference had some other problems :-/

Function subSetMarkText(oMark, sReplace As String)
    Dim oDupRange, nLen As Long, oRange, nNextParagraphBreak As Long, nLastParagraphBreak As Long
   
        subSetMarkText = oMark

        Set oRange = fnMarkRange(oMark)
        If oMark.supportsService("com.sun.star.text.Bookmark") Then
            Dim nReturnIndex, sParagraphBreak
            sParagraphBreak = Chr(13)

            ' Hack to include paragraph breaks in bookmarks
            oRange.String = fnReplace(sReplace, sParagraphBreak, Chr(10))

            nReturnIndex = InStr(sReplace, sParagraphBreak)
            While (nReturnIndex)
                oDupRange = oRange.Text.createTextCursorByRange(oRange)
                oDupRange.collapseToStart
                oDupRange.goRight(nReturnIndex-1, False)
                oDupRange.goRight(1, True)
                oRange.Text.insertControlCharacter(oDupRange, 0, True)
                nReturnIndex = InStr(nReturnIndex + 1, sReplace, sParagraphBreak)
            Wend
        ElseIf oMark.supportsService("com.sun.star.text.TextSection") Then
            oRange.String = sReplace
        ElseIf oMark.supportsService("com.sun.star.text.ReferenceMark") Then
            ' In OpenOffice 3.2, creating a SwXTextCursor from SwXTextRange
            ' and then changing the contents of SwXTextCursor deletes the anchor
            ' of the SwXTextRange, that it's needed later on
            ' Instead of it, inserts the new citation after the first character of the
            ' field (SwXTextRange), then deletes the first character, then deletes the
            ' old field.

            nLen = Len(oRange.String)
            oDupRange = oRange.Text.createTextCursorByRange(oRange)

            if nLen > 1 Then
                ' In this case it mantains the same field and changes the contents. See the comment
                ' of r28041 for more informattion

                ' Inserts the citation after the first character
                oDupRange.collapseToStart
                oDupRange.goRight(1, False)
                oDupRange.String = sReplace

                ' Remove the first character
                oDupRange = fnMarkRange(oMark)
                oDupRange = oDupRange.Text.createTextCursorByRange(oDupRange)
                oDupRange.collapseToStart
                oDupRange.goRight(1,True)
                oDupRange.String = ""

                ' Remove the last part (part that was originally here)
                oDupRange = fnMarkRange(oMark)
                oDupRange = oDupRange.Text.createTextCursorByRange(oDupRange)
                oDupRange.collapseToEnd
                oDupRange.goLeft(nLen - 1,True)
                oDupRange.String = ""
            Else
                    ' For more information about the bug that this code is workarounding:
                ' see Mendeley ticket #7938
                    ' In OpenOffice 3.2 I have not been able to edit the contents of one
                ' field that the total length is 1 (e.g. "Nature Journal" style for
                ' firsts entries. So it deletes the field and creates a new one
                ' that happens to be longer (contains the standard text added by
                ' fnAddMark and calls subSetMarkText again that will edit the citation
                ' with the above code (nLen > 1)
                Dim fieldName as String

                ' 14: is the random number that OpenOffice adds automatically
                ' to the field name
                fieldName = Left(oMark.Name, Len(oMark.Name) - 14)

                oDupRange.String = "" ' deletes the field

                Dim thisField
                subSetMarkText = fnAddMark(oDupRange, fieldName)
                subSetMarkText = subSetMarkText(subSetMarkText, sReplace)
            End If
        Else
            Print "shouldn't get here"
        End If
End Function

Function fnConvert(oMark)
    Dim oRange, sMarkName As String
    Dim markText As String
    
    Dim spaceAdded As Boolean
    spaceAdded = False
    
    
    markText = getMarkText(oMark)
    sMarkName = fnMarkName(oMark)
    Set oRange = fnMarkRange(oMark)
    
    
    Call subDeleteMark(oMark, True)
    Set fnConvert = fnAddMark(oRange, sMarkName)

        fnConvert = subSetMarkText(fnConvert, markText)
    
    
End Function


' Grabs bookmark name from document properties, if necessary
Function fnGetFullBookmarkName(sName As String) As String
    If (Left(sName, Len(ZOTERO_BOOKMARK_REFERENCE_PROPERTY)) = ZOTERO_BOOKMARK_REFERENCE_PROPERTY) Then
        ' Get property
        fnGetFullBookmarkName = fnGetProperty(sName)
    Else
        fnGetFullBookmarkName = sName
    End If
End Function

Function fnMarkName(oMark) As String
        fnMarkName = oMark.Name
        If oMark.supportsService("com.sun.star.text.Bookmark") Then
            fnMarkName = fnGetFullBookmarkName(fnMarkName)
        ElseIf Mid(fnMarkName, Len(fnMarkName)-REFERENCEMARK_RANDOM_STRING_LENGTH-Len(REFERENCEMARK_RANDOM_DATA_SEPARATOR)+1, 4) = REFERENCEMARK_RANDOM_DATA_SEPARATOR Then
          fnMarkName = Left(fnMarkName, Len(fnMarkName)-REFERENCEMARK_RANDOM_STRING_LENGTH-Len(REFERENCEMARK_RANDOM_DATA_SEPARATOR))
        End If
End Function

Function fnPrefix(ByVal sName As String) As String
    Dim nLoc As Long
    
    nLoc = InStr(sName, "MENDELEY_")
    fnPrefix = Mid(sName, nLoc, 11)
End Function

Function fnID(sName)
    Dim nLoc As Long
    
    nLoc = InStr(sName, "MENDELEY_")
    fnID = Mid(sName, nLoc + 12, Len(sName))
End Function

' Replaces a string with another
Function fnReplace(sString, sSearch, sReplace) As String
        Dim substrings
        substrings = Split(sString, sSearch)
        fnReplace = Join(substrings, sReplace)
End Function

Function fnOtherTextInNote(oRange) As Boolean
    'Dim nMarkLength As Long, nNoteLength As Long, oFootnote
    
        fnOtherTextInNote = Len(oRange.Text.String) > Len(oRange.String)
End Function

Sub subDeleteNote(oRange)
        oRange.Text.dispose
End Sub

Sub subDeleteMark(oMark, Optional bDontDeleteNote As Boolean)
    Dim oRange, oDupRange
    
    Set oRange = fnMarkRange(oMark)
    If Not bDontDeleteNote And fnLocationType(oRange) <> ZOTERO_MAIN Then
        If Not fnOtherTextInNote(oRange) Then
            Call subDeleteNote(oRange)
            Exit Sub
        End If
    End If
    
        Dim oVC, fnSelection
        ' Check to see if we need to delete the invisible character used to
        ' separate the reference mark from the user's text
        oDupRange = thisComponent.currentController.viewCursor.Text.createTextCursorByRange(oMark.Anchor)
        oDupRange.collapseToEnd
        oDupRange.goRight(1, True)
        If(oDupRange.String = Chr(0) Or oDupRange.String = Chr(8288)) Then  ' have invisible character
            oDupRange.String = ""
        End If

        If oMark.supportsService("com.sun.star.text.Bookmark") Then
            'Make sure any properties are gone
            Call subSetProperty(oMark.Name, "")

            oMark.Anchor.String = ""
            oMark.dispose
        ElseIf oMark.supportsService("com.sun.star.text.TextSection") Then
            oMark.Anchor.String = ""
            oMark.dispose
        ElseIf oMark.supportsService("com.sun.star.text.ReferenceMark") Then
            oMark.Anchor.String = ""
        End If
End Sub


Function isMendeleyCitationField(code As String) As Boolean
    isMendeleyCitationField = (startsWith(code, MENDELEY_CITATION) And Len(code) > Len(MENDELEY_CITATION)) _
        Or (startsWith(code, MENDELEY_EDITED_CITATION) And Len(code) > Len(MENDELEY_EDITED_CITATION)) _
        Or (startsWith(code, MENDELEY_CITATION_MAC) And Len(code) > Len(MENDELEY_CITATION_MAC)) _
        Or InStr(code, CSL_CITATION) > 0
End Function

Function isMendeleyBibliographyField(code As String) As Boolean
    isMendeleyBibliographyField = startsWith(code, MENDELEY_BIBLIOGRAPHY) Or startsWith(code, MENDELEY_BIBLIOGRAPHY_MAC) _
        Or InStr(code, CSL_BIBLIOGRAPHY) > 0
End Function

Function convertUuidsListToString(allUuidsList) As String
    ' Constructs a list with the UUIDS
    Dim first As Boolean
    first = True
    
    Dim UuidsFormatted As String
    UuidsFormatted = ""
    Dim uuid
    
    For Each uuid In allUuidsList
        If Not (uuid = "") Then
            If first = False Then
                UuidsFormatted = UuidsFormatted & ";"
            End If
            
            UuidsFormatted = UuidsFormatted & uuid
            first = False
        End If
    Next
    convertUuidsListToString = UuidsFormatted
End Function

Function activeDocumentPath() As String
  activeDocumentPath = thisComponent.URL
End Function
