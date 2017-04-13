' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009-2016 Mendeley Ltd.
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

' author: m.dhandapani@elsevier.com
Dim c_count as integer
Sub MovetoInline(ctxt1)
rem ----------------------------------------------------------------------
rem define variables
Dim document   as object
Dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
rem ----------------------------------------------------------------------
Dim args2(21) as new com.sun.star.beans.PropertyValue
args2(0).Name = "SearchItem.StyleFamily"
args2(0).Value = 2
args2(1).Name = "SearchItem.CellType"
args2(1).Value = 0
args2(2).Name = "SearchItem.RowDirection"
args2(2).Value = true
args2(3).Name = "SearchItem.AllTables"
args2(3).Value = false
args2(4).Name = "SearchItem.SearchFiltered"
args2(4).Value = false
args2(5).Name = "SearchItem.Backward"
args2(5).Value = false
args2(6).Name = "SearchItem.Pattern"
args2(6).Value = false
args2(7).Name = "SearchItem.Content"
args2(7).Value = false
args2(8).Name = "SearchItem.AsianOptions"
args2(8).Value = false
args2(9).Name = "SearchItem.AlgorithmType"
args2(9).Value = 0
args2(10).Name = "SearchItem.SearchFlags"
args2(10).Value = 0
args2(11).Name = "SearchItem.SearchString"
args2(11).Value = ""
args2(12).Name = "SearchItem.ReplaceString"
args2(12).Value = ctxt1
args2(13).Name = "SearchItem.Locale"
args2(13).Value = 255
args2(14).Name = "SearchItem.ChangedChars"
args2(14).Value = 2
args2(15).Name = "SearchItem.DeletedChars"
args2(15).Value = 2
args2(16).Name = "SearchItem.InsertedChars"
args2(16).Value = 2
args2(17).Name = "SearchItem.TransliterateFlags"
args2(17).Value = 256
args2(18).Name = "SearchItem.Command"
args2(18).Value = 2
args2(19).Name = "SearchItem.SearchFormatted"
args2(19).Value = false
args2(20).Name = "SearchItem.AlgorithmType2"
args2(20).Value = 1
args2(21).Name = "Quiet"
args2(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())
End Sub

Sub ExportAsBookmarks
rem ----------------------------------------------------------------------
rem define variables
Dim document   as object
Dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Delete", "", 0, Array())
End Sub

Sub Footnote_Inline()

 Dim xDoc,marks,fnotes, thisnote, nfnotes, fnotecount,c_count,c_limit
 Dim C_Mark()
 Dim mBookmarks
 Dim args1(0) as new com.sun.star.beans.PropertyValue
 Dim document   as object
 Dim dispatcher as object
 Dim j As Long
 
  
  xDoc = thiscomponent   
  fnotes=xdoc.getFootNotes()
  c_limit= fnotes.getCount()-1  
  
 On Error GoTo ErrorHandler
  	
 IF c_limit <= 0 Then
 	Exit sub
  Else
  	Redim C_Mark(c_limit)
 End If 
  
  If fnotes.hasElements() Then
  
    fnotecount=0
    For nfnotes=0 To fnotes.getCount()-1
	    thisnote=fnotes.getbyIndex(nfnotes)     
	    C_Mark(nfnotes) = thisnote.getString  
    Next nfnotes
    
  End If

     document   = ThisComponent.CurrentController.Frame
     dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
     mBookmarks = thisComponent.Bookmarks.ElementNames
  	 c_limit= fnotes.getCount()-1

	For i = 0 to c_limit	
		
		If C_Mark(i) ="" Then
			goto x:
		End If		
		Call Replace_Inline(C_Mark(i),mBookmarks(i))
		
	Next
X:	
	Call Remove_BookmarkAll
    Call refresh
    
GoTo EndOfSub
    ErrorHandler:
    Call reportError

    EndOfSub:
    uiDisabled = False
    
End Sub

Sub Replace_Inline(citationText,ccount)
rem ----------------------------------------------------------------------
rem define variables
Dim document   as object
Dim dispatcher as object
Dim oAnchor  'Bookmark anchor
Dim oCursor  'Cursor at the left most range.
Dim oMarks

rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
Dim args1(21) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.SearchFiltered"
args1(4).Value = false
args1(5).Name = "SearchItem.Backward"
args1(5).Value = false
args1(6).Name = "SearchItem.Pattern"
args1(6).Value = false
args1(7).Name = "SearchItem.Content"
args1(7).Value = false
args1(8).Name = "SearchItem.AsianOptions"
args1(8).Value = false
args1(9).Name = "SearchItem.AlgorithmType"
args1(9).Value = 0
args1(10).Name = "SearchItem.SearchFlags"
args1(10).Value = 0
args1(11).Name = "SearchItem.SearchString"
args1(11).Value = citationText
args1(12).Name = "SearchItem.ReplaceString"
args1(12).Value = ""
args1(13).Name = "SearchItem.Locale"
args1(13).Value = 255
args1(14).Name = "SearchItem.ChangedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.DeletedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.InsertedChars"
args1(16).Value = 2
args1(17).Name = "SearchItem.TransliterateFlags"
args1(17).Value = 256
args1(18).Name = "SearchItem.Command"
args1(18).Value = 0
args1(19).Name = "SearchItem.SearchFormatted"
args1(19).Value = false
args1(20).Name = "SearchItem.AlgorithmType2"
args1(20).Value = 1
args1(21).Name = "Quiet"
args1(21).Value = true
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
rem ----------------------------------------------------------------------
  oMarks = ThisComponent.getBookmarks()
  oAnchor = oMarks.getByName(ccount).getAnchor()
  oCursor = ThisComponent.getCurrentController().getViewCursor()
  oCursor.gotoRange(oAnchor, False)
  dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
  
End Sub


Sub Remove_BookmarkAll()
Dim mBookmarks
Dim j As Long
Dim args1(0) as new com.sun.star.beans.PropertyValue
Dim document   as object
Dim dispatcher as object

     document   = ThisComponent.CurrentController.Frame
     dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
     mBookmarks = thisComponent.Bookmarks.ElementNames
    For j = 0 To UBound(mBookmarks)
        
             args1(0).Name = "Bookmark"
             args1(0).Value = mBookmarks(j)
             dispatcher.executeDispatch(document, ".uno:DeleteBookmark", "", 0, args1())         
    Next
End Sub

'iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii

'Inline_To_Bookmark

Sub BKMK_Insert(Ct)
rem ----------------------------------------------------------------------
rem define variables
Dim document   as object
Dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
Dim args1(0) as new com.sun.star.beans.PropertyValue
args1(0).Name = "Bookmark"
args1(0).Value = Ct
dispatcher.executeDispatch(document, ".uno:InsertBookmark", "", 0, args1())
End Sub

'iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii
Sub Move_Inline_Footnote()

Dim marks,xDoc,fnotes,c_limit

On Error GoTo ErrorHandler
  xDoc = thiscomponent   
  fnotes=xdoc.getFootNotes()
  c_limit= fnotes.getCount()-1
   
	 IF c_limit <= 0 Then
	  Else
	  Call Footnote_Inline
	 End If 


	ZoteroUseBookmarks = false
    marks = fnGetMarks(ZoteroUseBookmarks)
    c_count = 0
    Dim mark
   
    For Each mark In marks
        citationText = getMarkText(mark)
        Call Replace_Footnote(citationText)        
    Next
    
Call refresh

GoTo EndOfSub
    ErrorHandler:
    Call reportError
    EndOfSub:
    uiDisabled = False
End Sub

Function Check_Bookmark(cTxt) as Boolean
     Dim mBookmarks
     Dim j As Long
     Dim args1(0) as new com.sun.star.beans.PropertyValue
     Dim document   as object
     Dim dispatcher as object

     document   = ThisComponent.CurrentController.Frame
     dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
     mBookmarks = thisComponent.Bookmarks.ElementNames
    For j = 0 To UBound(mBookmarks)
      
    msgbox  Left(cTxt, 13)
    msgbox Left(mBookmarks(j), 13)
    
    
        If Left(mBookmarks(j), 13) =  Left(cTxt, 13) Then
             args1(0).Name = "Bookmark"
            args1(0).Value = mBookmarks(j)
             
             Check_Bookmark = true
             Exit For
             
             'dispatcher.executeDispatch(document, ".uno:DeleteBookmark", "", 0, args1())
        End if
    Next
End Function

Sub Replace_Footnote(Ctxt)
rem ----------------------------------------------------------------------
rem define variables
Dim document   as object
Dim dispatcher as object
 
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
Dim args1(21) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.SearchFiltered"
args1(4).Value = false
args1(5).Name = "SearchItem.Backward"
args1(5).Value = false
args1(6).Name = "SearchItem.Pattern"
args1(6).Value = false
args1(7).Name = "SearchItem.Content"
args1(7).Value = false
args1(8).Name = "SearchItem.AsianOptions"
args1(8).Value = false
args1(9).Name = "SearchItem.AlgorithmType"
args1(9).Value = 0
args1(10).Name = "SearchItem.SearchFlags"
args1(10).Value = 0
args1(11).Name = "SearchItem.SearchString"
args1(11).Value = Ctxt
args1(12).Name = "SearchItem.ReplaceString"
args1(12).Value = ""
args1(13).Name = "SearchItem.Locale"
args1(13).Value = 255
args1(14).Name = "SearchItem.ChangedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.DeletedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.InsertedChars"
args1(16).Value = 2
args1(17).Name = "SearchItem.TransliterateFlags"
args1(17).Value = 256
args1(18).Name = "SearchItem.Command"
args1(18).Value = 0
args1(19).Name = "SearchItem.SearchFormatted"
args1(19).Value = false
args1(20).Name = "SearchItem.AlgorithmType2"
args1(20).Value = 1
args1(21).Name = "Quiet"
args1(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
Call BKMK_Insert(Ctxt)
rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())
rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertFootnote", "", 0, Array())
rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
End Sub


Function countCitationInFootnotes1() as integer
 	Dim oMarks,oMark,oTxt
	Dim Ftnt As Integer

    ZoteroUseBookmarks = True
    oMarks = fnGetMarks(ZoteroUseBookmarks)
    For Each oMark in oMarks
        oTxt= fnMarkRange(oMark)
        oTx1t=getMarkText(oMark)
    
        If fnLocationType(oTxt) = ZOTERO_FOOTNOTE Then        
        	msgbox oTx1t        
           Ftnt = Ftnt + 1	
        End If
   	Next
    countCitationInFootnotes1 = Ftnt
End Function

Function Footno_Validate(Cittext) as Boolean

 Dim marks
 Dim c_count 
 Dim xDoc
 Dim fnotes, thisnote, nfnotes, fnotecount
 dim c_limit
 Dim C_Mark()
  xDoc = thiscomponent   
  
  ' by popular demand ...
  fnotes=xdoc.getFootNotes()
   c_limit= fnotes.getCount()-1
   'msgbox c_limit
 IF c_limit <= 0 Then
  	Exit function
  Else
  	Redim C_Mark(c_limit)
 End If 
  
  If fnotes.hasElements() Then
    fnotecount=0
    For nfnotes=0 To fnotes.getCount()-1
	    thisnote=fnotes.getbyIndex(nfnotes)           
	    C_Mark(nfnotes) = thisnote.getString	 
	   'msgbox C_Mark(nfnotes) 	
	  	If Cittext = C_Mark(nfnotes) Then
	  	Footno = true
	  	End If 
  
    Next nfnotes
    
  End If


End Function


Function Footno() as boolean
 Dim xDoc
 Dim fnotes
 Dim c_limit
 Dim C_Mark()
	xDoc = thiscomponent   
	fnotes=xdoc.getFootNotes()
	c_limit= fnotes.getCount()-1
	
 If c_limit <= 0 Then
  	Exit function
  Else
  	Redim C_Mark(c_limit)
 End If    
End Function