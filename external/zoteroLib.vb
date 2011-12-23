
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

' All the following functions were originally copied from the Zotero code base
' (https://www.zotero.org/svn/integration/ice/trunk/plugin.vb - revision 3444)
' and have been modified to work properly with the rest of the Mendeley code

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

Function isStandardCslCitationField(code As String) As Boolean
    isStandardCslCitationField = InStr(code, CSL_CITATION) > 0
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
