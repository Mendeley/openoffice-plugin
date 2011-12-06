
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

Option Explicit

' DYMAMICARRAY:
' Tries to redim the array as less as possible but allows to grow up

Type DynamicArrayType
' In OO.org Basic Public Type is not available. So when we are declaring
' currentCitationsArray in mendeleyMain it uses Variant
    currentSize As Long
    ArrayContents As Variant
End Type

' DynamicStringType:
' This is a string which has extra spaces on the end, to fill up the allocated memory
' when appending, if there isn't enough allocated memory, double the allocated memory.
' This is more efficient than a normal string for lots of small string concatenations.

Type DynamicStringType
' In OO.org Basic Public Type is not available. So when we are declaring
' currentCitationsArray in mendeleyMain it uses Variant
    length As Long
    contents As String
End Type

Global Const ARRAY_INITIAL_SIZE = 10
Global Const ARRAY_INCREMENTS = 5

' ################ DynamicArray
Function DynamicArrayInit() As DynamicArrayType
    Dim ArrayContents() As String
    ReDim ArrayContents(0 To ARRAY_INITIAL_SIZE)
    
    Dim DynamicArray As DynamicArrayType
    DynamicArray.currentSize = 0
    DynamicArray.ArrayContents = ArrayContents
    DynamicArrayInit = DynamicArray
End Function

Function DynamicArrayAppend(DynamicArray As DynamicArrayType, value As String) As DynamicArrayType
    Dim lastPosition As Long
    lastPosition = DynamicArray.currentSize
       
    If lastPosition = UBound(DynamicArray.ArrayContents) Then
    Dim ArrayContents() as String
    ArrayContents = DynamicArray.ArrayContents

    Dim newSize as Long
    newSize = UBound(ArrayContents) + ARRAY_INCREMENTS
    ReDim Preserve ArrayContents(newSize)

    DynamicArray.ArrayContents = ArrayContents
    End If
    
    DynamicArray.ArrayContents(lastPosition) = value
    DynamicArray.currentSize = DynamicArray.currentSize + 1
    DynamicArrayAppend = DynamicArray
End Function

Function DynamicArrayGetPosition(DynamicArray As DynamicArrayType, position As Long) As String
    Dim value As String
    
    value = DynamicArray.ArrayContents(position)
    DynamicArrayGetPosition = value
End Function

Function DynamicArraySize(DynamicArray As DynamicArrayType) As Long
    DynamicArraySize = DynamicArray.currentSize
End Function

Function DynamicArrayHasElement(DynamicArray As DynamicArrayType, searchFor As String) As Boolean
    Dim i As Long
    DynamicArrayHasElement = False
    
    For i = 0 To UBound(DynamicArray.ArrayContents)
        If DynamicArray.ArrayContents(i) = searchFor Then
            DynamicArrayHasElement = True
            Exit Function
        End If
    Next
End Function

Function DynamicArrayTest()
    Dim DynamicArray As DynamicArrayType
    DynamicArray = DynamicArrayInit()
    
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test01")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test02")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test03")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test04")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test05")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test06")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test07")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test08")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test09")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test10")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test11")
    DynamicArray = DynamicArrayAppend(DynamicArray, "Test12")
    
    MsgBox DynamicArrayHasElement(DynamicArray, "dkdkdk") & " expected false"
    MsgBox DynamicArrayHasElement(DynamicArray, "Test01") & " expected true"
    MsgBox DynamicArrayHasElement(DynamicArray, "Test12") & " exepcted true"
    MsgBox DynamicArrayHasElement(DynamicArray, "Test07") & " expected true"
    
    MsgBox "ArraySize:" & DynamicArraySize(DynamicArray) & " expected 12"
    MsgBox "Position:" & DynamicArrayGetPosition(DynamicArray, 1) & " Test02"
End Function

' DynamicStringType functions:

Function DynamicStringInit(initialLength As Long) As DynamicStringType
    If initialLength < 1 Then
        initialLength = 1
    End If
    
    Dim stringContents As String
    stringContents = Space(initialLength)
    
    Dim dynamicString As DynamicStringType
    dynamicString.length = 0
    dynamicString.contents = stringContents
    
    DynamicStringInit = dynamicString
End Function

Function DynamicStringAppend(dynamicString As DynamicStringType, toAppend As String) As DynamicStringType
    If (Len(toAppend) > 0) Then
        ' if not enough allocated space in the string, keep doubling it till there is
        Do While dynamicString.length + Len(toAppend) > Len(dynamicString.contents)
            dynamicString.contents = dynamicString.contents & Space(Len(dynamicString.contents))
        Loop
        
        Mid$(dynamicString.contents, dynamicString.length + 1, Len(toAppend)) = toAppend
        dynamicString.length = dynamicString.length + Len(toAppend)
    End If
    
    DynamicStringAppend = dynamicString
End Function

Function DynamicStringGet(dynamicString As DynamicStringType) As String
    DynamicStringGet = Left$(dynamicString.contents, dynamicString.length)
End Function

Sub DynamicStringTest()
    Dim testString As DynamicStringType
    
    testString = DynamicStringInit(10)
    testString = DynamicStringAppend(testString, "Tim: Hello there!")
    MsgBox DynamicStringGet(testString)
    testString = DynamicStringAppend(testString, Chr(13) & "Tom: And hello to you!")
    MsgBox DynamicStringGet(testString)
    testString = DynamicStringAppend(testString, Chr(13) & "Tim: What an interesting conversation we are having!!!! :) ;-)")
    MsgBox DynamicStringGet(testString)
End Sub
