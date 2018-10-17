' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009-2012, 2017 Mendeley Ltd.
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

' These need to be Global as they are accessed from EventClassModule
Global previouslySelectedField
Global previouslySelectedFieldResultText As String

Global Const DEBUG_MODE = ${DEBUG_MODE}
Private Const DEBUG_LOG_FILE = "file:///F:/oo-debugLog.txt"

Global Const TEMPLATE_NAME_DURING_BUILD = "MendeleyPlugin"

Global Const MENDELEY_DOCUMENT = "Mendeley Document"
Global Const MENDELEY_CONTROL_BAR = "Mendeley Control Bar"
Global Const MENDELEY_USER_ACCOUNT = "Mendeley User Name"
Global Const MENDELEY_USER_UUID = "Mendeley Unique User Id"
Global Const INSERT_CITATION = "Insert Citation"
Global Const INSERT_BIBLIOGRAPHY = "Insert Bibliography"
Global Const MENDELEY_CITATION = "Mendeley Citation"
Global Const MENDELEY_CITATION_EDITOR = "Mendeley Citation Editor"
Global Const MENDELEY_CITATION_MAC = " PRINTDATE Mendeley Citation"
Global Const MENDELEY_EDITED_CITATION = "Mendeley Edited Citation"
Global Const MENDELEY_BIBLIOGRAPHY = "Mendeley Bibliography"
Global Const MENDELEY_BIBLIOGRAPHY_MAC = " PRINTDATE Mendeley Bibliography"
Global Const MENDELEY_CITATION_STYLE = "Mendeley Citation Style"
Global Const SELECT_ME_FETCH_STYLES = "Select me to fetch the styles"

Global Const CSL_CITATION = "CSL_CITATION "
Global Const CSL_BIBLIOGRAPHY = "CSL_BIBLIOGRAPHY"
Global Const CSL_BIBLIOGRAPHY_OLD = "CSL_BIBLIOGRAPHY "
Global Const INSERT_CITATION_TEXT = "{Formatting Citation}"
Global Const CITATION_EDIT_TEXT = "{Reformatting Citation}"
Global Const BIBLIOGRAPHY_TEXT = "{Bibliography}"
Global Const MERGING_TEXT = "{Merging Citations}"
Global Const TOOLBAR = "Mendeley Toolbar"
Global Const TOOLBAR_CITATION_STYLE = "Citation Style"
Global Const TOOLBAR_INSERT_CITATION = "Insert Citation"
Global Const TOOLBAR_EDIT_CITATION = "Edit Citation"
Global Const TOOLBAR_MERGE_CITATIONS = "Merge Citations"
Global Const TOOLBAR_INSERT_BIBLIOGRAPHY = "Insert Bibliography"
Global Const TOOLBAR_REFRESH = "Refresh"
Global Const TOOLBAR_EXPORT = "Export..."
Global Const TOOLBAR_UNDO_EDIT = "Undo Edit"
Global Const MERGE_CITATIONS_NOT_ENOUGH_CITATIONS = "Please select at least two citations before clicking 'Merge Citations'."
Global Const CITATIONS_NOT_ADJACENT = "Citations must be adjacent to merge."
Global Const CITATION_ADJECENT_LIMIT = 4
Global Const MACRO_ALREADY_RUNNING = "Waiting For Response From Mendeley Desktop"
Global Const DOCUMENT_NOT_IN_LIBRARY = "{Document not in library}"
Global Const RECENT_STYLE_NAME = "Mendeley Recent Style Name"
Global Const RECENT_STYLE_ID = "Mendeley Recent Style Id"

Global Const TOOLTIP_INSERT_CITATION = "Insert a new citation (Alt-M)"
Global Const TOOLTIP_EDIT_CITATION = "Edit the selected citation (Alt-M)"
Global Const TOOLTIP_UNDO_EDIT = "Undo custom edit of the selected citation"
Global Const TOOLTIP_MERGE_CITATIONS = "Merge the selected citations into one"
Global Const TOOLTIP_INSERT_BIBLIOGRAPHY = "Insert a bibliography"
Global Const TOOLTIP_REFRESH = "Refresh citations and bibliographies"
Global Const TOOLTIP_CITATION_STYLE = "Select a citation style"
Global Const TOOLTIP_EXPORT_OPENOFFICE = "Export a copy of the document compatible with LibreOffice"
Global Const TOOLTIP_EXPORT_WITHOUT_MENDELEY_FIELDS = "Export the document without Mendeley data fields"
Global Const TOOLTIP_EXPORT = "Export the document with different options"

Global Const CITATION_NUMBER = "citation-number"

Global Const COULDNT_OPEN_MENDELEY_MESSAGE = "Couldn't open Mendeley Desktop. Please run Mendeley Desktop first.\n\n(download the latest version from www.mendeley.com if necessary)"

' The following control how Zotero generates random ReferenceMark names in OOo
Global Const REFERENCEMARK_RANDOM_DATA_SEPARATOR = " RND"
Global Const REFERENCEMARK_RANDOM_STRING_LENGTH = 10

' The following dictate the maximum length of the strings to store in each of these data types
Global Const MAX_BOOKMARK_LENGTH = 50
Global Const MAX_PROPERTY_LENGTH = 255
Global Const BOOKMARK_ID_STRING_LENGTH = 10
Global Const ZOTERO_BOOKMARK_REFERENCE_PROPERTY = "Mendeley_Bookmark"

'The following constants describe a location in a document
Global Const ZOTERO_ERROR = 0 'Frame, comments, header, footer
Global Const ZOTERO_MAIN = 1 'Main document including things like tables (wdMainTextStory)
Global Const ZOTERO_FOOTNOTE = 2 'Footnote (wdFootnotesStory)
Global Const ZOTERO_ENDNOTE = 3 'Endnote (wdEndnotesStory)
Global Const ZOTERO_TABLE = 4 ' Inside a Table

' This is the conversion from twips to 100th mms, according to Google Calculator
Global Const TWIPS_TO_100TH_MM = 1.76388889

Global Const MAX_UUIDS = 2000

Global Const MSGBOX_RESULT_YES = 6

' Possible types for MsgBox (see https://help.libreoffice.org/Basic/MsgBox_Function_Runtime)
' Buttons and icons can be combined
' e.g. MSGBOX_TYPE_OK_CANCEL + MSGBOX_TYPE_EXCLAMATION
Global Const MSGBOX_TYPE_OK = 0
Global Const MSGBOX_TYPE_OK_CANCEL = 1
Global Const MSGBOX_TYPE_ABORT_RETRY_IGNORE = 2
Global Const MSGBOX_TYPE_YES_NO_CANCEL = 3
Global Const MSGBOX_TYPE_YES_NO = 4
Global Const MSGBOX_TYPE_RETRY_CANCEL = 5
Global Const MSGBOX_TYPE_STOP = 16
Global Const MSGBOX_TYPE_QUESTION = 32
Global Const MSGBOX_TYPE_EXCLAMATION = 48
Global Const MSGBOX_TYPE_INFORMATION = 64
Global Const MSGBOX_TYPE_FIRST_BUTTON_DEFAULT = 128
Global Const MSGBOX_TYPE_SECOND_BUTTON_DEFAULT = 256
Global Const MSGBOX_TYPE_THIRD_BUTTON_DEFAULT = 512

Global debugInfo As String
Global updateCitationComboDone As Boolean 'Set to False by default

Global initialised As Boolean
Global hangingIndentLength As Long ' The size of the hanging indent in points

' to prevent user from performing actions while we're still in the middle of another one
Global uiDisabled As Boolean
Global awaitingResponseFromMD As Boolean
Global ZoteroUseBookmarks As Boolean
Global openingWordDoc As Boolean

Global seedGenerated As Boolean

Global unitTest As Boolean

Dim mendeleyApi

Global Const JSON_CSL_CITATION = "CSL_CITATION "
Global Const JSON_PREVIOUS = "MendeleyPrevious"
Global Const JSON_URL = "MendeleyUrl"
Global Const VALIDATE_INSERT_AREA = "Mendeley can not insert a citation or bibliography at this location." 'Validate inserting area 

' Next TEMP_BOOKMARK to keep the position of the cursor on 
Global Const TEMP_BOOKMARK_CURSOR_POSITION = "MendeleyTempCursorBookmark"
Global Const TEMP_BOOKMARK_CURSOR_POSITION_STYLE = "MendeleyTempCursorBookmark_Style"
Global Const FOOTNOTE_CITATIONS_MERGE = "Footnote citations can't be merged at this location."

' arguments can be a single String argument or an Array of argument Strings
Function mendeleyApiCall(functionName As String, Optional arguments) As String
    If IsEmpty(mendeleyApi) Then
        mendeleyApi = createUnoService("com.sun.star.task.MendeleyDesktopAPI")
    End If
        
    Dim mArgs(0 to 0) As New com.sun.star.beans.NamedValue
    mArgs(0).Name = "function name"
    mArgs(0).Value = functionName

    If Not IsMissing(arguments) Then
        If IsArray(arguments) Then
            ReDim Preserve mArgs(0 to UBound(arguments)) As New com.sun.star.beans.NamedValue 
            Dim arg
            For arg = 1 to UBound(arguments)
                mArgs(arg).Name = "argument"
                mArgs(arg).Value = arguments(arg)
            Next    
        Else
            ReDim Preserve mArgs(0 to 1) As New com.sun.star.beans.NamedValue 
            mArgs(1).Name = "argument"
            Dim argument As String
            argument = arguments
            mArgs(1).Value = argument
        End If
    End If
    
    Dim returnVal
    returnVal = mendeleyApi.Execute(mArgs)
    
    Dim resultLength
    resultLength = simpleMendeleyApiCall("previousResultLength")

    If (simpleMendeleyApiCall("previousSuccess") = "False") Then
        Dim message As String
        message = simpleMendeleyApiCall("previousErrorMessage")

        If (DEBUG_MODE) Then
            message = message + Chr(10) + simpleMendeleyApiCall("previousResponse")
        End If

        MsgBox message
    End If

    mendeleyApiCall = returnVal
End Function

' Only used by the "previousSuccess" and "previousErrorMessage" calls
Function simpleMendeleyApiCall(functionName As String) As String
    If IsEmpty(mendeleyApi) Then
        mendeleyApi = createUnoService("com.sun.star.task.MendeleyDesktopAPI")
    End If
        
    Dim mArgs(0 to 0) As New com.sun.star.beans.NamedValue
    mArgs(0).Name = "function name"
    mArgs(0).Value = functionName

    Dim returnVal
    returnVal = mendeleyApi.Execute(mArgs)

    simpleMendeleyApiCall = returnVal
End Function

Function apiResetCitations() As String
    apiResetCitations = mendeleyApiCall("resetCitations")
End Function
Function apiFormatCitationsAndBibliography() As String
    apiFormatCitationsAndBibliography = mendeleyApiCall("formatCitationsAndBibliography")
End Function
Function apiAddCitation(fieldCode As String, footnoteIndex as Integer) As String
    Dim arguments(1 to 2) As String
    arguments(1) = fieldCode
    arguments(2) = CStr(footnoteIndex)
    apiAddCitation = mendeleyApiCall("addCitationCluster", arguments)
End Function
Function apiAddFormattedCitation(displayedText As String) As String
    apiAddFormattedCitation = mendeleyApiCall("addFormattedCitation", displayedText)
End Function
Function apiGetCitationJson(index As Long) As String
    apiGetCitationJson = mendeleyApiCall("getCitationCluster", index)
End Function
Function apiGetFormattedCitation(index As Long) As String
    apiGetFormattedCitation = mendeleyApiCall("getFormattedCitation", index)
End Function
Function apiGetFormattedBibliography() As String
    apiGetFormattedBibliography = mendeleyApiCall("getFormattedBibliography")
End Function
Function apiGetCitationStyleId() As String
    apiGetCitationStyleId = mendeleyApiCall("getCitationStyleId")
End Function
Function apiSetCitationStyle(styleId As String) As String
    apiSetCitationStyle = mendeleyApiCall("setCitationStyle", styleId)
End Function
Function apiGetUserAccount() As String
    apiGetUserAccount = mendeleyApiCall("getUserAccount")
End Function
Function apiGetUserUuid() As String
    apiGetUserUuid = mendeleyApiCall("getUserUuid")
End Function
Function apiGetCitationStylePresentationType() as Integer
    apiGetCitationStylePresentationType = mendeleyApiCall("getCitationStylePresentationType")
End Function
Function apiGetCitationStyleFromDialogServerSide(styleId As String) As String
    apiGetCitationStyleFromDialogServerSide = mendeleyApiCall("citationStyle_choose_interactive", styleId)
End Function
Function apiTestGetFieldCode(uuid As String) As String
    apiTestGetFieldCode = mendeleyApiCall("getFieldCodeFromUuid", uuid)
End Function
Function apiSetWordProcessor(wordProcessor As String, version As String) As String
    Dim args(1 to 2) As String
    args(1) = wordProcessor
    args(2) = version
    apiSetWordProcessor = mendeleyApiCall("wordProcessor_set", args)
End Function
Function apiUndoManualFormat(fieldCode As String) As String
    apiUndoManualFormat = mendeleyApiCall("citation_undoManualFormat", fieldCode)
End Function
Function apiCitationChoose(hintText As String) As String
    apiCitationChoose = mendeleyApiCall("citation_choose_interactive", hintText)
End Function
Function apiCitationEdit(fieldCode As String, hintText As String) As String
    Dim args(1 to 2) As String
    args(1) = fieldCode
    args(2) = hintText
    apiCitationEdit = mendeleyApiCall("citation_edit_interactive", args)
End Function
Function apiUpdateFieldCode(ByVal fieldCode) As String
    ' Note: not currently used or tested - could be useful for checking
    ' for manual edits immediately after user edit as in WinWord
    apiUpdateFieldCode = mendeleyApiCall("checkManualFormatAndGetFieldCode", fieldCode)
End Function
Function apiMergeCitations(fieldCodes) As String
    Dim args(1 to (1 + UBound(fieldCodes) - LBound(fieldCodes)))
    Dim index As Long
    For index = LBound(fieldCodes) to UBound(fieldCodes)
        args(index - LBound(fieldCodes) + 1) = fieldCodes(index)
    Next
    apiMergeCitations = mendeleyApiCall("citations_merge", args)
End Function
Function apiBringPluginToForeground() As String
    ' Note: This doesn't work at the moment
    ' apiBringPluginToForeground = mendeleyApiCall("bringPluginToForeground")
End Function
Function apiConnected() As Boolean
    apiConnected = mendeleyApiCall("isMendeleyDesktopRunningStr") = "True"

    If apiConnected = False Then
        MsgBox "Please run Mendeley Desktop before using the plugin.", Title:="Couldn't Connect To Mendeley Desktop"
    End If
End Function
Function apiGetDesktopCitationStyleId() As String
    apiGetDesktopCitationStyleId = mendeleyApiCall("getDesktopSelectedStyleId")
End Function

' initialise on word startup and on new / open document
Public Sub AutoExec()
    If Not USE_RIBBON Then
        Call initialise
    End If
End Sub
Public Sub AutoNew()
    Call initialise
End Sub
Public Sub AutoOpen()
    If buildingPlugin() Then
        Exit Sub
    End If
    
    Call initialise
End Sub

Public Sub initialise()
    uiDisabled = True
    ThisDocument.Saved = True
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    ' enable the mendeley plugin ui
    uiDisabled = False
End Sub

Sub debugLog(message As String)
    If Not DEBUG_MODE Then
        Exit Sub
    End If

    Dim debugLogFile
    debugLogFile = FreeFile
    Open DEBUG_LOG_FILE For Append As #debugLogFile
    Write #debugLogFile, "Time: " & Time(), message
    Close #debugLogFile
End Sub

' ----- Top level functions - those directly triggered by user actions -----
Sub mergeCitations()
     If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    Dim thisField As Variant
    '''''''''''''''''''''''''
    Dim oSelection As Object
    Dim oViewCursor As Object
    Dim selectionToReplace As Object
    Dim validatelocation As Integer
    Dim strSelectionCharCount As Integer
    Dim count as Long
    Dim fieldCodesToMerge(0) As String
    Dim presentationType as Integer

    oSelection = thisComponent.currentController.getViewCursor()
    validatelocation = fnLocationType(oSelection)
    'Validation for empty string and Location
    If validatelocation = ZOTERO_TABLE Then    	'Selected citation in the Table area
        oViewCursor = thiscomponent.getCurrentController().getViewCursor()
        selectionToReplace = oViewCursor.cell.createTextCursorByRange(oViewCursor.cell)
        strSelectionCharCount  = Len(selectionToReplace.String)
    ElseIf validatelocation = ZOTERO_MAIN Then
        strSelectionCharCount  =Len(oSelection.string)
    ElseIf Validatelocation = ZOTERO_FOOTNOTE Then
	    Goto Foot
    End If
     If validatelocation = ZOTERO_ERROR Then    'Validate the selected citation location
        MsgBox CITATIONS_NOT_ADJACENT, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citation"
        GoTo EndOfSub
     ElseIf strSelectionCharCount = 0 Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citations"
        GoTo EndOfSub
     ElseIf validatelocation = ZOTERO_FOOTNOTE Then 'To avoid merge citation in footnote area
        MsgBox FOOTNOTE_CITATIONS_MERGE, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citation"
        GoTo EndOfSub
     End If

     'Validate Footnote Merge


    presentationType = apiGetCitationStylePresentationType()
    If presentationType = ZOTERO_FOOTNOTE Then
        Dim strLen as Integer
        Dim strSelection as String
        Dim i as Integer
        Dim vVar As Variant
        Dim footNoteValue, Footn, markCode
        Dim footnoteText as String
        Dim footnoteCount as Integer
        Dim countEt as Integer
        Dim countTxt as Integer

        If validatelocation = ZOTERO_TABLE Then 'for table based citation string
            strSelection = selectionToReplace.string
        Else
            strSelection = oSelection.string
        End If
        strLen = Len(strSelection)
        set footNoteValue =  ThisComponent.getfootnotes()
        count = 0
        countEt = 0
        countTxt = 0

        For i = 0 to strLen-1
        vVar = Left(Right(strSelection,strLen-i),1)
            If  Asc ((Left(Right(strSelection,strLen-i),1))) = 13 then  'Validate ASCII value for enter
               countEt = countEt + 1
            ElseIf Asc(Left(Right(strSelection,strLen-i),1)) = 32 then   'Validate ASCII value for Space
               countEt = countEt + 1
            ElseIf IsNumeric(vVar) = True then
               vVar = Int(vVar) - 1
               Footn = footNoteValue.getByIndex(vVar)
               footnoteText =   Footn.getString()
               ReDim Preserve fieldCodesToMerge(0 to count) As String
               markCode = fnGetFieldCode(footnoteText)

               If isMendeleyCitationField(markCode) = False Then
                    MsgBox CITATIONS_NOT_ADJACENT
                    GoTo EndOfSub
               End If
               fieldCodesToMerge(count) = markCode
               count = count + 1
             ElseIf  IsNumeric(vVar) = False then
                 If  vVar = Chr(0) Or vVar = Chr(8288) Or vVar = Chr(10) Then
             Else
                 countTxt = countTxt +1
             End If
          End If
        Next

       If Count <= 1 Or countEt >= 5 Or countTxt <> 0 then
           MsgBox CITATIONS_NOT_ADJACENT
           GoTo EndOfSub
       Else
        Goto SkipField_Footnote
       End If
     End If

   'merge normal citation (not footnote) included table's citation
Foot:
    Dim allDocumentMarks, mark
    Dim markName
    Dim fieldRange
    Dim lastPossibleStartOfNextField
    Dim spl_condition

    allDocumentMarks = getSelectedCitationMarks(False) 'Store only selected citation
    If IsEmpty(allDocumentMarks) = true Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citations"
        GoTo EndOfSub
    ElseIf UBound(allDocumentMarks) = 0 Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citations"
        GoTo EndOfSub
    End If

    If Not apiConnected() Then
        GoTo EndOfSub
    End If

    For Each mark In allDocumentMarks
        Set thisField = mark
        fieldRange = fnMarkRange(thisField)
        markName = getMarkName(thisField)
        If isMendeleyCitationField(markName) = False Then
            GoTo SkipField
        End If

    'Valdation for space or text between citations also table citation 
       If count = 0 Then
            If  Validatelocation = ZOTERO_FOOTNOTE  then
                oViewCursor = thiscomponent.getCurrentController().getViewCursor()
                selectionToReplace = oViewCursor.getText.createTextCursorByRange(fieldRange.getStart())
            ElseIf validatelocation = ZOTERO_TABLE Then
                selectionToReplace = oViewCursor.cell.createTextCursorByRange(oViewCursor.cell.getstart())
            Else
                selectionToReplace = thisComponent.Text.createTextCursorByRange(fieldRange.getStart())
            End If
       Else

            If validatelocation = ZOTERO_TABLE Then
                    lastPossibleStartOfNextField = oViewCursor.cell.createTextCursorByRange(fieldRange.getStart())
                    lastPossibleStartOfNextField.goLeft(CITATION_ADJECENT_LIMIT,False)
            ElseIf Validatelocation = ZOTERO_FOOTNOTE Then
                   lastPossibleStartOfNextField = oViewCursor.getText.createTextCursorByRange(fieldRange.getStart())
            Else
                    lastPossibleStartOfNextField = thisComponent.Text.createTextCursorByRange(fieldRange.getStart())
            End If
            
            lastPossibleStartOfNextField.goLeft(CITATION_ADJECENT_LIMIT,False)
            
            If Validatelocation = ZOTERO_FOOTNOTE Then
                If oViewCursor.getText.compareRegionEnds(selectionToReplace,lastPossibleStartOfNextField) = 1 Then
                    MsgBox CITATIONS_NOT_ADJACENT
                    GoTo EndOfSub
                EndIf
            Else
               If thisComponent.Text.compareRegionEnds(selectionToReplace,lastPossibleStartOfNextField) = 1 Then
                   MsgBox CITATIONS_NOT_ADJACENT
                    GoTo EndOfSub
                EndIf
            End If

       End If

       If validatelocation = ZOTERO_TABLE Then
            While oViewCursor.Cell.compareRegionEnds(fieldRange, selectionToReplace) = -1
                selectionToReplace.goRight(1, True)
            WEnd
      Else

        If Validatelocation = ZOTERO_FOOTNOTE Then
            while oViewCursor.getText.compareRegionEnds(fieldRange, selectionToReplace) = -1
            selectionToReplace.goRight(1, True)
                WEnd
        Else
            While thisComponent.Text.compareRegionEnds(fieldRange, selectionToReplace) = -1
                    selectionToReplace.goRight(1, True)
                WEnd
        End If
      End If

        ReDim Preserve fieldCodesToMerge(0 to count) As String
        fieldCodesToMerge(count) = markName
        count = count + 1

SkipField:

    Next
    If count < 2 Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Merge Citations"
      GoTo EndOfSub
    End If

    '''''''''''''''''''''''''''''''''''''''''''
SkipField_Footnote:

    Dim newFieldCode As String
    newFieldCode = apiMergeCitations(fieldCodesToMerge)
    Dim citeField
    If presentationType = ZOTERO_FOOTNOTE Then
        If validatelocation = ZOTERO_TABLE Then
            Set citeField = fnAddMark(selectionToReplace,newFieldCode,"")
        Else
            Set citeField = fnAddMark(oSelection,newFieldCode,"")
        End If
    Else
        If validatelocation = ZOTERO_TABLE Then
            Set citeField = fnAddMark(selectionToReplace,newFieldCode,"")
        Else
            Set citeField = fnAddMark(oSelection,newFieldCode,"")
        End If

    End IF

    Call RefreshDocument

    GoTo EndOfSub
    ErrorHandler:
    Call reportError
    EndOfSub:
    uiDisabled = False
End Sub

Sub gotoBookmark(bkmk as String)
    Dim oAnchor  'Bookmark anchor
    Dim oCursor  'Cursor at the left most range.
    Dim oMarks

On Error GoTo ErrorHandler

    oMarks = ThisComponent.getBookmarks()
    oAnchor = oMarks.getByName(bkmk).getAnchor()
    oCursor = ThisComponent.getCurrentController().getViewCursor()
    oCursor.gotoRange(oAnchor, False)

ErrorHandler:

End Sub

Sub insertBookmark(bookmarkValue As String)
    Dim document as Object
    Dim dispatcher as Object
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    Dim args1(0) as new com.sun.star.beans.PropertyValue
    args1(0).Name = "Bookmark"
    args1(0).Value = bookmarkValue
    dispatcher.executeDispatch(document, ".uno:InsertBookmark", "", 0, args1())
End Sub

Sub deleteBookmark(bookmarkValue as String)
    Dim document   as Object
    Dim dispatcher as Object
On Error GoTo ErrorHandler
    rem get access to the document
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    Dim args1(0) as new com.sun.star.beans.PropertyValue
    args1(0).Name = "Bookmark"
    args1(0).Value = bookmarkValue
    dispatcher.executeDispatch(document, ".uno:DeleteBookmark", "", 0, args1())
ErrorHandler:
End Sub

Sub privateInsertCitation(hintText As String)
    Dim currentMark
    Dim citeField
    
    Dim bringToForeground As Boolean
    bringToForeground = False
    'Insert the cursor bookmark to keep cursor in original position
    Call insertbookmark(TEMP_BOOKMARK_CURSOR_POSITION)
    'Validate Insert area
    Dim validateLocation
    validateLocation = thisComponent.currentController.viewCursor
    If fnLocationType(validateLocation) = ZOTERO_ERROR Then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Citation"
        GoTo EndOfSub
    End If

    Set citeField = Nothing
    
    currentMark = getFieldAtSelection()
    
    Dim markName As String
    
    If Not (currentMark Is Nothing) And Not IsEmpty(currentMark) Then
        Dim fieldType As String
    
        markName = getMarkName(currentMark)
        
        If isMendeleyCitationField(markName) Then
            ' fine
        ElseIf isMendeleyBibliographyField(markName) Then
            MsgBox "Bibliographies are generated automatically and cannot be manually edited"
            GoTo EndOfSub
        Else
            MsgBox "This is not an editable citation."
            GoTo EndOfSub
        End If
    End If
    
    If Not apiConnected() Then
        GoTo EndOfSub
    End If
    
    Call sendWordProcessorVersion
    
    Dim useCitationEditor As Boolean
    useCitationEditor = True
    
    ZoteroUseBookmarks = False

    Dim selectedRange
    Set selectedRange = fnSelection()
    If (selectedRange Is Nothing) Then Return
    
    Dim citationText As String
    If useCitationEditor Then
        citationText = CITATION_EDIT_TEXT
        citeField = getFieldAtSelection()

        Dim currentSelection
        currentSelection = fnSelection()

        Dim position
    Else
        citationText = MENDELEY_CITATION
    End If
    
    If (citeField Is Nothing) Or IsEmpty(citeField) Then
    End If

    Call setMendeleyDocument(True)
        
    If apiConnected() Then
        If Not isDocumentLinkedToCurrentUser Then
            GoTo EndOfSub
        End If
    
        Dim stringLength As Long
        awaitingResponseFromMD = True
        
        Dim fieldCode As String
        If unitTest Then
            fieldCode = apiTestGetFieldCode(hintText)
        Else
            If Len(markName) = 0 Then
                fieldCode = apiCitationChoose(markName, hintText)
            Else
                fieldCode = apiCitationEdit(markName, hintText)
            End If
        End If

        awaitingResponseFromMD = False
        
        If (Right(fieldCode, Len("DoNotBringToForeground")) = "DoNotBringToForeground") Then
            fieldCode = Left(fieldCode, Len(fieldCode) - Len("DoNotBringToForeground"))
            bringToForeground = False
        Else
            bringToForeground = True
            Call apiBringPluginToForeground
        End If
         
        ' check for null result:
        If (Len(fieldCode) = 0) Or ((Len(fieldCode) = 1) And (fieldCode = "")) Then
            ' MsgBox "No Citation Received from Mendeley"
            GoTo EndOfSub
        Else
              ' Need to add a space after range before insert, so typing does not
              ' extend field
              selectedRange.collapseToEnd
              selectedRange.String = Chr(8288) ' Zero-width non-joiner
              Dim oDupRange
              oDupRange = selectedRange.Text.createTextCursorByRange(selectedRange)
              selectedRange.collapseToStart
            If (currentMark Is Nothing) Or IsEmpty(currentMark) Then
                If (CLng(citationText) > 65535) Then
                    MsgBox "Error: Response too long." + Chr(10) + Chr(10) + "Please don't cite so many references in one citation."
                End If

                'Check Style type and insert footnote
                Dim presentationType as Integer
                Dim intlastFootnote as Integer
                Dim document as Object
                Dim dispatcher as Object
                Dim Omrk

                presentationType = apiGetCitationStylePresentationType()
                If presentationType = ZOTERO_FOOTNOTE Then
                    document   = ThisComponent.CurrentController.Frame
                    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
                    dispatcher.executeDispatch(document, ".uno:InsertFootnote", "", 0, array())
                    Omrk =  ThisComponent.getCurrentController().getViewCursor()
                    Set citeField = fnAddMark(Omrk, citationText, "")
                Else
                    Set citeField = fnAddMark(selectedRange, citationText, "")
                End If
            Else
                  citeField = currentMark
                  citeField = subSetMarkText(citeField, citationText)
            End If
        End If
        
        ' check if another instance of Word is awaiting a response from Mendeley
        If fieldCode = "<CURRENTLY-PROCESSING-REQUEST>" Then
            MsgBox "You can only make one citation at a time, " + _
                "please choose the documents for your initial citation in Mendeley Desktop first"
            GoTo EndOfSub
        End If
        Set citeField = fnRenameMark(citeField, fieldCode)
        Call refreshDocument(False)
    End If
    'Move cursor to original position
    Call gotoBookmark(TEMP_BOOKMARK_CURSOR_POSITION)
    'Delete inserted bookmark
    call deleteBookmark(TEMP_BOOKMARK_CURSOR_POSITION)
    GoTo EndOfSub

ErrorHandler:
    Call reportError
    
EndOfSub:
    If Not (citeField Is Nothing) And Not IsEmpty(citeField) Then
      If Not (citeField.Anchor Is Nothing) Then
        If getMarkText(citeField) = INSERT_CITATION_TEXT Or _
            getMarkText(citeField) = CITATION_EDIT_TEXT Then
                Call subDeleteMark(citeField)
        End If
      End If
    End If
    
    If bringToForeground Then
        Call apiBringPluginToForeground
    End If
End Sub

Sub insertBibliography()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    'Validate Insert area
    Dim validateLocation
    validateLocation = thisComponent.currentController.viewCursor
    If fnLocationType(validateLocation) = ZOTERO_ERROR Then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Bibliography"
        GoTo EndOfSub
    End If

    If isCursorInBibliography() = True then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Citation or Bibliography"
        Goto EndOfSub
    End If
    

    ZoteroUseBookmarks = False
    
    Dim fieldAtSelection As Variant
    fieldAtSelection = Nothing
    fieldAtSelection = getFieldAtSelection()

    If Not(fieldAtSelection Is Nothing) Then
        If isObject(fieldAtSelection) Then
            MsgBox "A bibliography cannot be inserted within another citation or bibliography."
            GoTo EndOfSub
        End If
    End If
    
    Call setMendeleyDocument(True)
    
    If Not apiConnected() Then
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If

    'Insert space before and after from Bibliography to move cursor from section.
    Call moveCursorFromSection
    
    Dim thisField 'As Field
    Set thisField = fnAddMark(fnSelection(), "ADDIN " & MENDELEY_BIBLIOGRAPHY & " " & CSL_BIBLIOGRAPHY_OLD,"")
    
    Call refreshDocument
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    uiDisabled = False
End Sub
Sub moveCursorFromSection()
    Dim document   as object
    Dim dispatcher as object
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    Dim args1(0) as new com.sun.star.beans.PropertyValue
    'Insert space
    args1(0).Name = "Text"
    args1(0).Value = "  "
    dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args1())
    dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, Array())

End Sub

Function isCursorInBibliography() as Boolean
   Dim currentSelection as object
   Dim cursorSelectionString as String
On Error goto ErrorHandler
   ' ErrorHandler becasue getting the ViewCursor in certain places (e.g. in a field)
   ' returns an error.
   currentSelection = thisComponent.currentController.getViewCursor()
   cursorSelectionString = currentSelection.Textsection.name
   isCursorInBibliography = isMendeleyBibliographyField(cursorSelectionString)
   Exit Function
ErrorHandler:
     isCursorInBibliography = False
End Function
Sub undoEdit()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Const NOT_IN_EDITABLE_CITATION_TITLE = "Undo Citation Edit"
    Const NOT_IN_EDITABLE_CITATION_TEXT = "Place the cursor within an edited citation before clicking 'Undo Edit'"
    
    If isCursorInBibliography() = True then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Citation or Bibliography"
        Goto EndOfSub
    End If

    Dim currentMark
    
    If Not IsEmpty(getFieldAtSelection()) Then
        Set currentMark = getFieldAtSelection()
    End If
    
    If currentMark Is Nothing Or IsEmpty(currentMark) Then
        MsgBox NOT_IN_EDITABLE_CITATION_TEXT, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, NOT_IN_EDITABLE_CITATION_TITLE
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If
    
    Dim markName As String
    markName = getMarkName(currentMark)
        
    Dim newMarkName As String
    newMarkName = apiUndoManualFormat(markName)
    
    currentMark = fnRenameMark(currentMark, newMarkName)
    currentMark = subSetMarkText(currentMark, INSERT_CITATION_TEXT)
    
    Call refreshDocument
    GoTo EndOfSub
    
ErrorHandler:
   Call reportError
    
EndOfSub:
   uiDisabled = False
End Sub

Sub refresh()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True

    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub Remove_Bookmark()
     Dim mBookmarks
     Dim j As Long
     Dim args1(0) as new com.sun.star.beans.PropertyValue
     Dim document   as object
     Dim dispatcher as object

     document   = ThisComponent.CurrentController.Frame
     dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
     mBookmarks = thisComponent.Bookmarks.ElementNames
    For j = 0 To UBound(mBookmarks)
        If (Left(mBookmarks(j), 9) = "Mendeley_" Or InStr(mBookmarks(j), "CSL_CITATION") > 0 Or InStr(mBookmarks(j), "CSL_BIBLIOGRAPHY") > 0) Then
             args1(0).Name = "Bookmark"
             args1(0).Value = mBookmarks(j)
             dispatcher.executeDispatch(document, ".uno:DeleteBookmark", "", 0, args1())
         End if
    Next
End Sub

Sub chooseCitationStyle()
    If isUiDisabled() Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Dim chosenStyle As String
    
    Call setMendeleyDocument(True)
        
    If apiConnected() Then
        chosenStyle = apiGetCitationStyleFromDialogServerSide(getCitationStyleId())
        Call setCitationStyle(chosenStyle)
        Call refreshDocument
    End If
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub


Sub afterSave()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub afterOpen()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub insertCitationButton()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
        
    Dim tipText As String
    tipText = "Tip: You can press Alt-M instead of clicking Insert Citation."
    If isCursorInBibliography() = True then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Citation or Bibliography"
        Goto EndOfSub
    End If
    Call privateInsertCitation(tipText)
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub insertCitation()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True

     If isCursorInBibliography() = True then
        MsgBox VALIDATE_INSERT_AREA, MSGBOX_TYPE_OK + MSGBOX_TYPE_EXCLAMATION, "Insert Citation or Bibliography"
        Goto EndOfSub
    End If
    Call privateInsertCitation("")
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub exportCompatibleMSWord()
    dim oFileDialog
    dim sFilePickerArgs

    sFilePickerArgs = Array(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_AUTOEXTENSION)
    oFileDialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")

    oFileDialog.Initialize(sFilePickerArgs())
    oFileDialog.setTitle("Save To...")
    oFileDialog.appendFilter("Microsoft Word (*.doc)", "*.doc")
    oFileDialog.appendFilter("All Files", "*.*")

    dim sFileUrl
    if oFileDialog.execute() then
        dim sFiles

        sFiles = oFileDialog.Files
        sFileUrl = sFiles(0)        ' get the first file
        'MsgBox(sFileUrl)            ' show it
    else
        sFileUrl=""
    end if

    oFileDialog.Dispose()
    Call showWarningExportMSWordWithFootnotes
    Call exportAsBookmarks(sFileUrl)
    call Remove_Bookmark
End Sub

Sub showWarningExportMSWordWithFootnotes()
    'To identify the citations in Footnote area ad pop up the warning message while export to MS-WORD
    If countCitationInFootnotes() <> 0 Then
        Msgbox ("If you export a document as a .doc format and have inserted a citation in a footnote then you may experience an error.")
    End if 
End Sub

Function countCitationInFootnotes() as Integer
    Dim oMarks,oMark,oTxt
    Dim Ftnt As Integer

    ZoteroUseBookmarks = True
    oMarks = fnGetMarks(ZoteroUseBookmarks)
    For Each oMark in oMarks
        oTxt= fnMarkRange(oMark)
        If fnLocationType(oTxt) = ZOTERO_FOOTNOTE Then
           Ftnt = Ftnt + 1	
        End If
   	Next
    countCitationInFootnotes = Ftnt
End Function

Sub exportAsBookmarks(fileUrl)
    If fileUrl <> "" Then
        Dim marks
        ZoteroUseBookmarks = True
        marks = fnGetMarks(ZoteroUseBookmarks)

        dim exportProperties(1) as new com.sun.star.beans.PropertyValue
        exportProperties(0).Name = "FilterName"
        ThisComponent.storeToUrl(fileUrl, exportProperties())
    End If
End Sub

' ----- end of top level functions -----
Function isUiDisabled() As Boolean
    If uiDisabled Then
        Dim userResponse
        If awaitingResponseFromMD Then
            userResponse = MsgBox ("Please finish selecting a citation from Mendeley Desktop before continuing.", _
            MB_RETRYCANCEL, MACRO_ALREADY_RUNNING)

            If userResponse = IDRETRY Then
                If awaitingResponseFromMD Then
                    userResponse = MsgBox ("If you're not in the middle of selecting a citation something has gone wrong (sorry)." _
                        + Chr(13) + Chr(13) + "Do you want to re-enable the toolbar?", _
                        MB_YESNO, "Re-Enable Mendeley Toolbar And Continue?")
                End If
            End If
        Else
            userResponse = MsgBox ("The Mendeley toolbar is currently inactive because another macro hasn't finished running." + _
                Chr(13) + Chr(13) + "Do you want to re-activate it?", 4, "Re-Enable Mendeley Toolbar?")
        End If

        Select Case userResponse
            Case IDYES
                uiDisabled = False
                awaitingResponseFromMD = False
            Case IDRETRY
                uiDisabled = False
                awaitingResponseFromMD = False
        End Select
    End If
    isUiDisabled = uiDisabled
End Function

