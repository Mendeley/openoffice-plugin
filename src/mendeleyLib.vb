' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009-2017 Mendeley Ltd.
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
      Call apiSetWordProcessor("LibreOffice.org", "unknown")
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
          "Mendeley LibreOffice Plugin"
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

    If Not apiConnected() Then
        Exit Function
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        Exit Function
    End If
    
    Call sendWordProcessorVersion
    
    Dim oldCitationStyle As String
    oldCitationStyle = getCitationStyleId()
    ' Ensure Mendeley Desktop knows which citation style to use and that it's added to
    ' document properties
    Call setCitationStyle(oldCitationStyle)
    
    ' Subscribe to events (e.g. WindowSelectionChange) doing on refreshDocument as it
    ' doesn't work in initialise() when addExternalFunctions() is also called
    Dim citationNumberCount As Long
    citationNumberCount = 0
    
    Dim bibliography As String

    Call apiResetCitations
    
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
            Call apiAddCitation(addUnicodeTags(markName))
            
            ' Just send an empty string if the displayed text is a temporary placeholder
            Dim displayedText As String
            displayedText = getMarkText(thisField)
            'displayedText = getMarkTextWithFormattingTags(thisField)
            If displayedText = INSERT_CITATION_TEXT _
                    Or displayedText = MERGING_TEXT _
                    Or displayedText = CITATION_EDIT_TEXT Then
                displayedText = ""
            End If
            Call apiAddFormattedCitation(addUnicodeTags(displayedText))
            
        End If
    Next
    
    ' Now that we've compiled the list of uuids, give them to Mendeley Desktop
    ' and tell it to format the citations and bibliography
    Dim status As String
    status = apiFormatCitationsAndBibliography
        
    If Left(status, Len("failed")) = "failed" Then
        Call apiBringPluginToForeground
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
			fieldText = apiGetFormattedCitation(citationNumber)

			' no longer used
            'Dim previousFormattedCitation As String
            'previousFormattedCitation = formatUnicode(apiGetPreviouslyFormattedCitation(citationNumber))
            
            Dim jsonData As String
            jsonData = apiGetCitationJson(citationNumber)

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
                bibliography = bibliography + apiGetFormattedBibliography()
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

                ' Remove the initial paragraph new line.
                oDupRange = oRange.Text.createTextCursorByRange(oRange)

                oDupRange.goLeft(0,False)
                oDupRange.goRight(1,True)
                oDupRange.String = ""
        End If
        
        If Not (fieldText = "") Then
            ' Put text in field
        End If
        
    Next
    
    If Not unitTest Then
        Dim newCitationStyle As String
        newCitationStyle = apiGetCitationStyleId()
        
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
    Dim presentationType as Integer
    presentationType = apiGetCitationStylePresentationType()
    If presentationType = ZOTERO_FOOTNOTE Then
        Call convertInlineToFootnote
    Else
        Call convertFootnote_Inline
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
    Call apiSetCitationStyle(style)
End Sub

Function getCitationStyleId() As String
    getCitationStyleId = fnGetProperty(MENDELEY_CITATION_STYLE)
    If getCitationStyleId = "" Then
        getCitationStyleId = getDesktopCitationStyleId()
        Call subSetProperty(MENDELEY_CITATION_STYLE, getCitationStyleId)
    End If
End Function

Function getDesktopCitationStyleId() As String
    getDesktopCitationStyleId = apiGetDesktopCitationStyleId()
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

' returns the user unique Id which this document is currently linked to
Function mendeleyUserUuid() As String
    On Error GoTo CatchError

    mendeleyUserUuid = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_UUID).value
    Exit Function

CatchError:
    mendeleyUserUuid = ""
End Function

Sub setMendeleyUserAccount(value As String)
    On Error GoTo CatchError

    Dim test As String

    ' if MENDELEY_DOCUMENT property not set this will throw an exception
    test = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value

    ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value = value

    Exit Sub
CatchError:
    ActiveDocument.CustomDocumentProperties.Add(Name:=MENDELEY_USER_ACCOUNT, _
    LinkToContent:=False, value:=value, Type:=msoPropertyTypeString)
End Sub

Sub setMendeleyUserUuid(value() As Byte)
    On Error GoTo CatchError

    Dim test As String

    ' if MENDELEY_DOCUMENT property not set this will throw an exception
    test = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_UUID).value

    ActiveDocument.CustomDocumentProperties(MENDELEY_USER_UUID).value = value

    Exit Sub
CatchError:
    ActiveDocument.CustomDocumentProperties.Add(name:=MENDELEY_USER_UUID, _
        LinkToContent:=False, value:=value, Type:=msoPropertyTypeString)
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
    Dim currentMendeleyUuid As String
    Dim thisDocumentUuid As String
    Dim currentMendeleyUser As String
    Dim thisDocumentUser As String

    currentMendeleyUser = apiGetUserAccount()
    currentMendeleyUuid = apiGetUserUuid()
    thisDocumentUser = fnGetProperty(MENDELEY_USER_ACCOUNT)
    thisDocumentUuid = fnGetProperty(MENDELEY_USER_UUID)

    ' remove server protocol from account string
    thisDocumentUser = Replace(thisDocumentUser, "http://", "")
    thisDocumentUser = Replace(thisDocumentUser, "https://", "")

    If thisDocumentUser = "" And thisDocumentUuid = currentMendeleyUuid Then
        isDocumentLinkedToCurrentUser = True
    ElseIf currentMendeleyUser = thisDocumentUser Then
        isDocumentLinkedToCurrentUser = True
        Call subSetProperty(MENDELEY_USER_ACCOUNT, "")
        Call subSetProperty(MENDELEY_USER_UUID, currentMendeleyUuid)
    Else
        Dim result ' As VbMsgBoxResult
        Dim vbCrLf
        vbCrLf = Chr(13)

        If thisDocumentUuid = "" And thisDocumentUser = "" Then
            ' if no user currently linked then set without asking user
            result = MSGBOX_RESULT_YES
        Else
            ' ask user if they want to link the document to their account
            result = MsgBox("This document has been edited by another Mendeley user" + vbCrLf + vbCrLf + _
                "Do you wish to enable the Mendeley plugin to edit the citations and bibliography yourself?" + vbCrLf + vbCrLf, _
                MSGBOX_TYPE_YES_NO, "Enable Mendeley plugin for this document?")
        End If

        If result = MSGBOX_RESULT_YES Then
            Call subSetProperty(MENDELEY_USER_ACCOUNT, "")
            Call subSetProperty(MENDELEY_USER_UUID, currentMendeleyUuid)
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
    ' <span style="font-variant:small-caps;"></span> small-caps

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
    
    Call applyStyleToSpan(subRange, startPosition, endPosition)
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

Sub applyStyleToSpan(wholeRange as range, startPosition as Long, endPosition as Long)
    Dim startTag As String
    Dim endTag As String
    
	' Added condition becuase of avoid multiple execute the function for tag
    If InStr(rangeString(wholeRange), "<span style=""baseline"">") <> 0 Then
        startTag = "<span style=""baseline"">"
        endTag = "</span>"
        Call applyStyleToTags(startTag, endTag, wholeRange, startPosition, endPosition, "baseline")
    End If

    If InStr(rangeString(wholeRange), "<span style=""font-variant:small-caps;"">") <> 0 Then
        startTag = "<span style=""font-variant:small-caps;"">"
        endTag = "</span>"
        Call applyStyleToTags(startTag, endTag, wholeRange, startPosition, endPosition, "small-caps")
    End If
End Sub

Sub applyStyleToTagPairs(tag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long)
    Call applyStyleToTags("<" + tag + ">", "</" + tag + ">", wholeRange, startPosition, endPosition, tag)
End Sub

Sub applyStyleToTags(startTag As String, endTag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long, operation As String)
    Dim thisRange 'As Range
    Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
    
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
        thisRange.goRight(len(startTag), True)
        thisRange.String = ""
        
        ' find and remove end tag
        
        Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)

        Dim endTagPosition As Long
        
            endTagPosition = InStr(fnReplace(rangeString(thisRange), Chr(13), ""), endTag) - 1
            thisRange.goRight(endTagPosition, False)
            thisRange.goRight(len(endTag), True)
            ' add a space on the end, which is deleted later. Otherwise, if we are
            ' at the end of a mark, new characters will be inserted OUTSIDE the mark
            thisRange.String = " "

            Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
            thisRange.goRight(startTagPosition, False)
            thisRange.goRight(endTagPosition - startTagPosition, True)
        
        ' apply style
        Select Case operation
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
                
            Case "small-caps"
            	thisRange.setPropertyValue("CharCaseMap", 4)

            Case "baseline"
                thisRange.setPropertyValue("CharEscapement", 0)
                thisRange.setPropertyValue("CharEscapementHeight", 100)

        End Select

        ' remove extra space from after the tag
        Call thisRange.collapseToEnd()
        thisRange.goRight(1, True)
        thisRange.String = ""

        Set thisRange = wholeRange.Text.createTextCursorByRange(wholeRange)
    Loop
    
    If startTag = "<second-field-align>" And maxFirstFieldLength > 0 Then
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
    ' Note: this isn't currently used in the OpenOffice plugin
    ' it would be useful if we check for citation edits immediately
    ' after the user's cursor leaves a citation field as for the
    ' WinWord plugin

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
            
            If Not apiConnected() Then
                ' Don't need do anything - this edit will get detected next time the user refreshes
                Exit Sub
            End If
            
            Dim markName As String
            markName = getMarkName(previouslySelectedField)
            
            Dim displayedText As String
            displayedText = addUnicodeTags(getMarkText(previouslySelectedField))
            displayedText = Replace(displayedText, Chr(13), "<p>")
            
            Dim newMarkName As String
            newMarkName = apiUpdateCitation(markName, displayedText)
            
            If markName <> newMarkName Then
                Call fnRenameMark(previouslySelectedField, newMarkName)
                needUpdate = True
            End If
        End If
    End If
    
    If needUpdate Then
        Call refreshDocument
    End If
End Sub

' Replaces oMark with the other mark format e.g. com.sun.star.text.Bookmark or com.sun.star.text.ReferenceMark
' This is used to export documents into Ms Word format or from Ms Word into LibreOffice
Function ChangeMarkFormat(oMark, fieldType as String)
    Dim oRange, oNewMark
    Dim citationText as String, citationCode as String

    oRange = oMark.Anchor

    citationText = getMarkText(oMark)
    citationCode = fnMarkName(oMark)

    oNewMark = thisComponent.createInstance(fieldType)
    oNewMark.setName (INSERT_CITATION_TEXT)
    
    If fieldType = "com.sun.star.text.Bookmark" Then
        deleteInvisibleCharacter(oRange)
    End If
    
    oRange.text.insertTextContent(oRange, oNewMark, True)
    oNewMark = fnRenameMark(oNewMark, citationCode)
   
    ChangeMarkFormat = oNewMark
End Function

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

Sub convertInlineToFootnote()
    Dim marks      
    Dim markName As String
    Dim thisField
    Dim mark
    Dim markTxt As String 
    Dim oCursor, oField
    
    'Insert bookmark when style has changed to note
    Call insertBookmark(TEMP_BOOKMARK_CURSOR_POSITION_STYLE)
    marks = fnGetMarks(ZoteroUseBookmarks)
    For Each mark In marks
        Set thisField = mark
        markName = getMarkName(thisField)
        markTxt = getMarkText(thisField)
        If insertFootnote_DeleteCitation(thisField) = True Then
            oCursor = ThisComponent.getCurrentController().getViewCursor()
            Set oField = fnAddMark(oCursor,markName, markTxt)
        End If
    Next
    'Move to inital position
    Call gotoBookmark(TEMP_BOOKMARK_CURSOR_POSITION_STYLE)
    'Delete inserted temp bookmark
    Call deleteBookmark(TEMP_BOOKMARK_CURSOR_POSITION_STYLE)
End sub
 
Sub insertFootnote_DeleteCitation(oMark) as Boolean
    Dim ocur, oRange
    Dim document As object
    Dim dispatcher As object
    Dim oCursor
    
    Set oRange = fnMarkRange(oMark)
    If oMark.supportsService("com.sun.star.text.ReferenceMark") Then
        oCursor = ThisComponent.getCurrentController().getViewCursor()
      
        oCur = oMark.Anchor
        If fnLocationType(oCur) = ZOTERO_FOOTNOTE Then
            insertFootnote_DeleteCitation = False
        Exit Sub
        Else
            insertFootnote_DeleteCitation = True
        End If

        oMark.Anchor
        oMark.Anchor.String = ""

        oCursor.gotoRange(oCur,False)
        document = ThisComponent.CurrentController.Frame
        dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
        dispatcher.executeDispatch(document, ".uno:InsertFootnote", "", 0, Array())
    End If
End Sub


Function fnGetFieldCode(strTxt as String) as String
    Dim markName As String
    Dim markTxt As String
    Dim matchFlag As Integer
    Dim marks, thisField
    Dim mark
    Dim Omrk
    Dim oField
    Dim i As Integer
    
    marks = fnGetMarks(ZoteroUseBookmarks)
    matchFlag = 0
    i = 0
    For Each mark In marks
        Set thisField = mark
        markName = getMarkName(mark)
        markTxt = getMarkText(mark)
        If Trim(markTxt) = Trim(strTxt) Then
            matchFlag = 1
            fnGetFieldCode = markName
            Exit For
        End If
    Next
    If matchFlag = 0 Then
        fnGetFieldCode = ""
    End If
End Function

Sub convertFootnote_Inline()
    Dim foots
    Dim footn
    Dim markCode, footnoteText, oField, oCursor, Omrk
    Dim i as Integer
    Dim j as Integer
    
    i = 0
    j = 0
    
    Set foots = ThisComponent.getFootnotes()
    If foots.getCount() = 0 Then
        Exit Sub
    End If

    For i = 0 To foots.getCount() - 1
        footn = foots.getByIndex(j)
        Omrk = foots.getByIndex(j).getAnchor()
        oCursor = ThisComponent.getCurrentController().getViewCursor()
        footnoteText = footn.getString()
        If  Left(footn.string,1) = Chr(0) Or Left(footn.string,1) = Chr(8288) Then
            footnoteText  =  Right(footnoteText,Len(footnoteText)-1)
        End If
        If Right(footn.string,1)= Chr(0) Or Right(footn.string,1) = Chr(8288) Then
            footnoteText = Left(footnoteText,Len(footnoteText)-1)
        End If
        markCode = fnGetFieldCode(footnoteText)
        If markCode <> "" Then
            Set oField = fnAddMark(Omrk, markCode, footnoteText)
        End If
    Next 
End sub
