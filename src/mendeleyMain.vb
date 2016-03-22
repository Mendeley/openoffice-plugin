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
Function apiAddCitation(fieldCode As String) As String
    apiAddCitation = mendeleyApiCall("addCitationCluster", fieldCode)
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

    Dim allDocumentMarks
    allDocumentMarks = fnGetMarks(False)

    Dim thisField

    '''''''''''''''''''''''''
    Dim oSelection
    oSelection = thisComponent.currentController.getViewCursor()
    '''''''''''''''''''''''''

    ' The number of citation fields selected to merge
    Dim count as Long
    count = 0

    Dim mark
    Dim uuids As String
    Dim uuid

    Dim markName
    Dim fieldRange

    Dim selectionToReplace
    Dim lastPossibleStartOfNextField

    If Not apiConnected() Then
        GoTo EndOfSub
    End If

    Dim fieldCodesToMerge(0) As String

    For Each mark In allDocumentMarks
        Set thisField = mark

        fieldRange = fnMarkRange(thisField)

        If Not(thisComponent.Text.compareRegionStarts(oSelection,fieldRange.end) = 1 _
            And thisComponent.Text.compareRegionEnds(fieldRange.start,oSelection) = 1) Then
                Goto SkipField
        End If

        markName = getMarkName(thisField)

        If isMendeleyCitationField(markName) = False Then
            GoTo SkipField
        End If

        If count = 0 Then
            selectionToReplace = thisComponent.Text.createTextCursorByRange(fieldRange.getStart())
        Else
            ' if more than CITATION_ADJECENT_LIMIT characters gap between fields show error
            lastPossibleStartOfNextField = thisComponent.Text.createTextCursorByRange(fieldRange.getStart())
            lastPossibleStartOfNextField.goLeft(CITATION_ADJECENT_LIMIT,False)

            If thisComponent.Text.compareRegionEnds(selectionToReplace,lastPossibleStartOfNextField) = 1 Then
                MsgBox CITATIONS_NOT_ADJACENT
                GoTo EndOfSub
            EndIf
        EndIf

        '(would be good to avoid looping but I don't know how to figure out number of steps in advance)
        While thisComponent.Text.compareRegionEnds(fieldRange, selectionToReplace) = -1
            selectionToReplace.goRight(1, True)
        WEnd

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

    Dim newFieldCode As String
    newFieldCode = apiMergeCitations(fieldCodesToMerge)

    Dim citeField
    Set citeField = fnAddMark(selectionToReplace,newFieldCode)
    Call RefreshDocument

    GoTo EndOfSub
    ErrorHandler:
    Call reportError

    EndOfSub:
    uiDisabled = False
End Sub

Sub privateInsertCitation(hintText As String)
    Dim currentMark
    
    Dim bringToForeground As Boolean
    bringToForeground = False
    
    Dim citeField
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
    
        Dim buttonText As String
            buttonText = "Send Citation to\nLibreOffice Writer;Cancel\nCitation"
        
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

                Set citeField = fnAddMark(selectedRange, citationText)
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
    
    Dim thisField 'As Field
    Set thisField = fnAddMark(fnSelection(), "ADDIN " & MENDELEY_BIBLIOGRAPHY & " " & CSL_BIBLIOGRAPHY_OLD)
    
    Call refreshDocument
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    uiDisabled = False
End Sub

Sub undoEdit()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Const NOT_IN_EDITABLE_CITATION_TITLE = "Undo Citation Edit"
    Const NOT_IN_EDITABLE_CITATION_TEXT = "Place the cursor within an edited citation before clicking 'Undo Edit'"
    
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
    
    Call exportAsBookmarks(sFileUrl)
End Sub

Sub exportAsBookmarks(fileUrl)
    If fileUrl <> "" Then
        Dim marks
        ZoteroUseBookmarks = True
        marks = fnGetMarks(ZoteroUseBookmarks)

        dim exportProperties(1) as new com.sun.star.beans.PropertyValue
        exportProperties(0).Name = "FilterName"
        exportProperties(0).Value = "MS Word 97"
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

