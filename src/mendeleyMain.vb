
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


' These need to be Global as they are accessed from EventClassModule
Global previouslySelectedField
Global previouslySelectedFieldResultText As String

Global Const DEBUG_MODE = ${DEBUG_MODE}
Private Const DEBUG_LOG_FILE = "file:///D:/tasks/OODebug/debugLog.txt"

Global Const TEMPLATE_NAME_DURING_BUILD = "MendeleyPlugin"

Global Const MENDELEY_DOCUMENT = "Mendeley Document"
Global Const MENDELEY_CONTROL_BAR = "Mendeley Control Bar"
Global Const MENDELEY_USER_ACCOUNT = "Mendeley User Name"
Global Const INSERT_CITATION = "Insert Citation"
Global Const INSERT_BIBLIOGRAPHY = "Insert Bibliography"
Global Const MENDELEY_CITATION = "Mendeley Citation"
Global Const MENDELEY_CITATION_EDITOR = "Mendeley Citation Editor"
Global Const MENDELEY_CITATION_MAC = " PRINTDATE Mendeley Citation"
Global Const MENDELEY_EDITED_CITATION = "Mendeley Edited Citation"
Global Const MENDELEY_BIBLIOGRAPHY = "Mendeley Bibliography"
Global Const MENDELEY_BIBLIOGRAPHY_MAC = " PRINTDATE Mendeley Bibliography"
Global Const MENDELEY_CITATION_STYLE = "Mendeley Citation Style"
Global Const DEFAULT_CITATION_STYLE = "http://www.zotero.org/styles/apa"
Global Const DEFAULT_CITATION_STYLE_NAME = "American Psychological Association"
Global Const SELECT_ME_FETCH_STYLES = "Select me to fetch the styles"

Global Const CSL_CITATION = "CSL_CITATION "
Global Const CSL_BIBLIOGRAPHY = "CSL_BIBLIOGRAPHY "
Global Const INSERT_CITATION_TEXT = "{Formatting Citation}"
Global Const CITATION_EDIT_TEXT = ""
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
Global Const MERGE_CITATIONS_NOT_ENOUGH_CITATIONS = "Please select at least two citations to merge."
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
Global Const TOOLTIP_EXPORT_OPENOFFICE = "Export a copy of the document compatible with OpenOffice"
Global Const TOOLTIP_EXPORT_WITHOUT_MENDELEY_FIELDS = "Export the document without Mendeley data fields"
Global Const TOOLTIP_EXPORT = "Export the document with different options"

Global Const CONNECTION_CONNECTED = 0
Global Const CONNECTION_VERSION_MISMATCH = 1
Global Const CONNECTION_MENDELEY_DESKTOP_NOT_FOUND = 2
Global Const CONNECTION_NOT_A_MENDELEY_DOCUMENT = 3

Global Const MENDELEY_RPC_CONNECTION_FAILED = "MendeleyRpcConnectionFailed"

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
Global Const MSGBOX_BUTTONS_YES_NO = 4

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

Function testFunc()

	'mendeleyApiCall("setNumberTest", "5 4 3 2 1")
	'MsgBox mendeleyApiCall("getNumberTest", "")
	'Exit Function
	apiResetCitations()
	apiSetCitationStyle("http://www.zotero.org/styles/apa")
	apiAddCitation("Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
	apiAddFormattedCitation("formattedCitation1")
	
	mendeleyApiCall("formatCitationsAndBibliography", "")
	
	MsgBox "citation 1: " + mendeleyApiCall("getCitationCluster", "0")
	MsgBox "formatted citation 1: " + mendeleyApiCall("getFormattedCitation", "0")
End Function

Function mendeleyApiCall(functionName As String, argument As String) As String
	If IsEmpty(mendeleyApi) Then
		mendeleyApi = createUnoService("com.sun.star.task.MendeleyDesktopAPI")
	End If
	
	Dim mArgs(0 to 1) As New com.sun.star.beans.NamedValue
	mArgs(0).Name = "function name"
	mArgs(0).Value = functionName
	mArgs(1).Name = "argument"
	mArgs(1).Value = argument
	
	Dim returnVal
	returnVal = mendeleyApi.Execute(mArgs)
	'MsgBox "from API: " + returnVal
	mendeleyApiCall = returnVal
End Function
Function apiResetCitations() As String
    apiResetCitations = mendeleyApiCall("resetCitations", "")
End Function
Function apiFormatCitationsAndBibliography() As String
    apiFormatCitationsAndBibliography = mendeleyApiCall("formatCitationsAndBibliography", "")
End Function
Function apiAddCitation(ByVal fieldCode As String) As String
    apiAddCitation = mendeleyApiCall("addCitationCluster", fieldCode)
End Function
Function apiAddFormattedCitation(ByVal displayedText As String) As String
    apiAddFormattedCitation = mendeleyApiCall("addFormattedCitation", displayedText)
End Function
Function apiGetCitationJson(ByVal index As Long) As String
    apiGetCitationJson = mendeleyApiCall("getCitationCluster", index)
End Function
Function apiGetFormattedCitation(ByVal index As Long) As String
    apiGetFormattedCitation = mendeleyApiCall("getFormattedCitation", index)
End Function
Function apiGetFormattedBibliography() As String
    apiGetFormattedBibliography = mendeleyApiCall("getFormattedBibliography", "")
End Function
'Function apiGetPreviouslyFormattedCitation(ByVal index As Long) As String
'    apiGetPreviouslyFormattedCitation = mendeleyRpcCall("getPreviouslyFormattedCitation", index)
'End Function
Function apiGetCitationStyleId() As String
    apiGetCitationStyleId = mendeleyApiCall("getCitationStyleId", "")
End Function
Function apiSetCitationStyle(ByVal styleId As String) As String
    apiSetCitationStyle = mendeleyApiCall("setCitationStyle", styleId)
End Function

' - HTTP requests instead of linking to the LinkToMendeleyVba2.dll for OpenOffice
Function mendeleyRpcCall(functionName As String, argument As String, optional quitOnError As Boolean) As String
    Dim mendeleyRpc
    If IsEmpty(quitOnError) Then
        quitOnError = true
    End If
    mendeleyRpc = createUnoService("com.sun.star.task.MendeleyRPC")
    Dim mArgs(0) As New com.sun.star.beans.NamedValue
    mArgs(0).Name = "meaningless"
    mArgs(0).Value = functionName + Chr(13) + argument
    On Error Goto ErrorHandler
    mendeleyRpcCall = mendeleyRpc.Execute(mArgs)
    Exit Function
ErrorHandler:
    mendeleyRpcCall = MENDELEY_RPC_CONNECTION_FAILED
    If quitOnError Then
        MsgBox "Connection to Mendeley lost"
        End
    End If
End Function
Function extGetCitationUuidsFromDialog (ByVal buttonText As String) As Long
    extGetCitationUuidsFromDialog = mendeleyRpcCall("getCitationUuidsFromDialog", buttonText)
End Function
Function privateExtGetStringResult (ByRef result As String) As Long
    result = mendeleyRpcCall("getStringResult", result)
End Function
Function extCheckConnectionAndSetupIfNeeded() As Long
    Dim rpcResult As String
    rpcResult = mendeleyRpcCall("checkConnectionAndSetupIfNeeded", "", false)
    If (rpcResult = MENDELEY_RPC_CONNECTION_FAILED) Then
        extCheckConnectionAndSetupIfNeeded = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND
    Else
        extCheckConnectionAndSetupIfNeeded = rpcResult
    End If
End Function
Function extGetCitationStyleNames() As Long
    extGetCitationStyleNames = mendeleyRpcCall("getCitationStyleNames", "")
End Function
Function extGetCitationStyle() As Long
    extGetCitationStyle = mendeleyRpcCall("getCitationStyle", "")
End Function
Function extLaunchMendeley() As Long
    ' Don't know how to launch mendeley without linking to dll so present info to user instead
    MsgBox "Please run Mendeley Desktop before using the plugin.", Title:="Couldn't Connect To Mendeley Desktop"
    extLaunchMendeley = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND
End Function
Function extGetUserAccount() As Long
    extGetUserAccount = mendeleyRpcCall("getUserAccount", "")
End Function
Function extGetCitationStyleFromDialogServerSide(ByVal styleId As String) As Long
    extGetCitationStyleFromDialogServerSide = mendeleyRpcCall("getCitationStyleFromDialogServerSide", styleId)
End Function
Function extBringPluginToForeground() As Long
    extBringPluginToForeground = mendeleyRpcCall("bringPluginToForeground", "")
End Function
Function extGetFieldCodeFromCitationEditor(ByVal uuids As String) As Long
    extGetFieldCodeFromCitationEditor = mendeleyRpcCall("getFieldCodeFromCitationEditor", uuids)
End Function
Function extStartMerge() As Long
    extStartMerge = mendeleyRpcCall("startMerge", "")
End Function
Function extAddFieldCodeToMerge(ByVal fieldCodeToMerge As String) As Long
    extAddFieldCodeToMerge = mendeleyRpcCall("addFieldCodeToMerge", fieldCodeToMerge)
End Function
Function extGetMergedFieldCode() As Long
    extGetMergedFieldCode = mendeleyRpcCall("getMergedFieldCode", "")
End Function
Function extSetDisplayedText(ByVal displayedText) As Long
    extSetDisplayedText = mendeleyRpcCall("setDisplayedText", displayedText)
End Function
Function extCheckManualFormatAndGetFieldCode(ByVal fieldCode) As Long
    extCheckManualFormatAndGetFieldCode = mendeleyRpcCall("checkManualFormatAndGetFieldCode", fieldCode)
End Function
Function extGetDisplayedText() As Long
    extGetDisplayedText = mendeleyRpcCall("getDisplayedText", "")
End Function
Function extUndoManualFormat(ByVal fieldCode) As Long
    extUndoManualFormat = mendeleyRpcCall("undoManualFormat", fieldCode)
End Function
Function extSetWordProcessor(ByVal wordProcessor) As Long
    extSetWordProcessor = mendeleyRpcCall("setWordProcessor", wordProcessor)
End Function
Function extTestGetFieldCode(ByVal fieldCode) As Long
    extTestGetFieldCode = mendeleyRpcCall("testGetFieldCode", fieldCode)
End Function

' Allocates a string of the required length and calls extGetStringResult() to fill
' it in with the result of the previous dll function call
Public Function extGetStringResult(stringLength As Long) As String
    Dim result As String
    If stringLength < 0 Then
        MsgBox "Connection with Mendeley Desktop lost. Please restart Mendeley Desktop and try again."
        End
    End If
    If stringLength = 0 Then
        extGetStringResult = result
        Exit Function
    End If
    ' ensure sufficent space in string
    If stringLength >= 65535 Then
        MsgBox "Mendeley is trying to send a string of " & stringLength & " characters which exceeds the limit of 65535" & Chr(13) & _
            "Please don't cite so many documents in the same in-line citation."
        End
    End If
    ' ensure sufficent space in string
    result = String(stringLength, " ")
    privateExtGetStringResult (result)
    extGetStringResult = result
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

    If Not (launchMendeleyIfNecessary() = CONNECTION_CONNECTED) Then
        GoTo EndOfSub
    End If

    Call extGetStringResult(extStartMerge)

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

        Call extGetStringResult(extAddFieldCodeToMerge(markName))

        count = count + 1

SkipField:

    Next
    If count < 2 Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS
      GoTo EndOfSub
    End If

    '''''''''''''''''''''''''''''''''''''''''''

    Dim newFieldCode As String
    newFieldCode = extGetStringResult(extGetMergedFieldCode())

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
    
    Dim connectionStatus As Long
    connectionStatus = extCheckConnectionAndSetupIfNeeded()
    If connectionStatus = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND Then
        MsgBox "Please run Mendeley Desktop before using the plugin"
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

        'position = currentSelection.getPosition()

        'thisComponent.getCaretPosition

    Else
        citationText = MENDELEY_CITATION
    End If
    
    If (citeField Is Nothing) Or IsEmpty(citeField) Then
    End If

    Call setMendeleyDocument(True)
        
    If connectionStatus = CONNECTION_CONNECTED Then
        If Not isDocumentLinkedToCurrentUser Then
            GoTo EndOfSub
        End If
    
        Dim buttonText As String
            buttonText = "Send Citation to\nOpenOffice Writer;Cancel\nCitation"
        
        Dim stringLength As Long
        awaitingResponseFromMD = True
        
        Dim hintAndFieldCode As String
        hintAndFieldCode = hintText & ";" & markName
        If unitTest Then
            stringLength = extTestGetFieldCode(hintAndFieldCode)
        Else
            stringLength = extGetFieldCodeFromCitationEditor(hintAndFieldCode)
        End If

        awaitingResponseFromMD = False
        Dim fieldCode As String
        fieldCode = extGetStringResult(stringLength)
        
        If (Right(fieldCode, Len("DoNotBringToForeground")) = "DoNotBringToForeground") Then
            fieldCode = Left(fieldCode, Len(fieldCode) - Len("DoNotBringToForeground"))
            bringToForeground = False
        Else
            bringToForeground = True
            ' bring Word to foreground
            ' (this call doesn't work reliably on all systems,
            '  on my Windows 7 dev machine it's fine but on my
            '  Win7 and WinXP VMs it works about 50% of the time)
            Call extGetStringResult(extBringPluginToForeground)
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
                Set citeField = fnAddMark(selectedRange, citationText)
            Else
                  citeField = currentMark
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
        ' bring Word to foreground
        '  (with limited testing this SEEMS to work all the time on all systems, presumably because
        '   word has finished running the macro and is in a more responsive state.)
        Call extGetStringResult(extBringPluginToForeground)
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
    
    If Not (launchMendeleyIfNecessary() = CONNECTION_CONNECTED) Then
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If
    
    Dim thisField 'As Field
    Set thisField = fnAddMark(fnSelection(), "ADDIN " & MENDELEY_BIBLIOGRAPHY & " " & CSL_BIBLIOGRAPHY)
    
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
    Const NOT_IN_EDITABLE_CITATION_TEXT = "Place cursor within an edited citation and press this button to undo the edit"
    
    Dim currentMark
    
    If Not IsEmpty(getFieldAtSelection()) Then
        Set currentMark = getFieldAtSelection()
    End If
    
    If currentMark Is Nothing Or IsEmpty(currentMark) Then
        MsgBox NOT_IN_EDITABLE_CITATION_TEXT, 1, NOT_IN_EDITABLE_CITATION_TITLE
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If
    
    Dim markName As String
    markName = getMarkName(currentMark)
        
    Dim newMarkName As String
    newMarkName = extGetStringResult(extUndoManualFormat(markName))
    
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
        
    If launchMendeleyIfNecessary() = CONNECTION_CONNECTED Then
        chosenStyle = extGetStringResult(extGetCitationStyleFromDialogServerSide(getCitationStyleId()))
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