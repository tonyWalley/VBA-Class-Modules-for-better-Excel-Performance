VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: StrMgr
Option Explicit

Public Enum padType
    padLeft = 1
    padRight = 2
    
End Enum


Public Function PadAndTrim(value As String, padType As padType, valueLen As Integer, Optional padChr As String = " ") As String
    Dim currentLength As Integer
    currentLength = Len(value)
    
    If currentLength > valueLen Then
        ' Trim the string if it is longer than the specified length
        PadAndTrim = Left(value, valueLen)
    ElseIf currentLength < valueLen Then
        ' Pad the string if it is shorter than the specified length
        Select Case padType
            Case padLeft
                PadAndTrim = String(valueLen - currentLength, padChr) & value
            Case padRight
                PadAndTrim = value & String(valueLen - currentLength, padChr)
        End Select
    Else
        ' Return the string unchanged if it is already the correct length
        PadAndTrim = value
    End If
End Function
' Function to parse the string by a delimiter and trim each part
Public Function parseAndTrim(value As String, DELIMITER As String) As Variant
    Dim parts As Variant
    parts = Split(value, DELIMITER)
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)

        parts(i) = Trim(parts(i))
    Next i
    parseAndTrim = parts
End Function

' Public function to build a dynamic string based on a dictionary of replacement values
Public Function buildDynamicString(baseString As String, replaceValDic As Scripting.Dictionary) As String
    Dim key As Variant
    Dim regEx As Object
    Dim matchCollection As Object
    Dim forEachMatch As Object

    ' Regular expression object for finding placeholders and FOR_EACH blocks
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.pattern = "\[FOR_EACH:\d+\{[^}]+\}|\[\w+\]"

    ' Find all placeholders in the baseString
    Set matchCollection = regEx.Execute(baseString)

    ' Iterate through each match to replace placeholders with actual values
    For Each forEachMatch In matchCollection
        Dim pattern As String
        pattern = forEachMatch.value

        ' Check if the pattern is a FOR_EACH case
        If Left(pattern, 9) = "[FOR_EACH" Then
            ' Extract the number to identify the corresponding dictionary key
            Dim forEachKey As String
            forEachKey = "FOR_EACH_" & Mid(pattern, 11, InStr(pattern, "{") - 11)

            ' Retrieve the corresponding dictionary if it exists
            If replaceValDic.Exists(forEachKey) Then
                ' Remove the FOR_EACH wrapper and process the inner content
                Dim innerContent As String
                innerContent = Mid(pattern, InStr(pattern, "{") + 1, Len(pattern) - InStr(pattern, "{") - 2)
                baseString = Replace(baseString, pattern, getForEachString(innerContent, replaceValDic(forEachKey)))
            End If
        Else
            ' Strip the brackets and replace the placeholder
            pattern = Replace(pattern, "[", "")
            pattern = Replace(pattern, "]", "")
            If replaceValDic.Exists(pattern) Then
                baseString = Replace(baseString, "[" & pattern & "]", replaceValDic(pattern))
            End If
        End If
    Next forEachMatch

    ' Return the modified string
    buildDynamicString = baseString
End Function

' Helper function to process FOR_EACH blocks
Private Function getForEachString(forEachString As String, forEachDic As Scripting.Dictionary) As String
    Dim resultString As String
    Dim itemKey As Variant
    Dim tempString As String
    
    ' Iterate through each dictionary entry to replace placeholders in the template
    For Each itemKey In forEachDic
        tempString = forEachString ' Start with a fresh copy of the forEach template string
        tempString = buildDynamicString(tempString, forEachDic(itemKey))
        resultString = resultString & tempString & vbCrLf
    Next itemKey

    getForEachString = resultString
End Function

' Function to convert a number to an Excel-style column letter
Public Function convertNumToAlpha(N As Long) As String
    Dim result As String
    Dim remainder As Long
    
    ' Handle numbers less than or equal to 26
    If N <= 26 Then
        convertNumToAlpha = Chr(N + 64) ' A=1, B=2, etc.
        Exit Function
    End If

    ' For numbers greater than 26, use recursion
    Do While N > 0
        N = N - 1 ' Adjust for 0-based indexing in Excel column naming
        remainder = N Mod 26
        result = Chr(65 + remainder) & result
        N = N \ 26 ' Integer division by 26
    Loop

    convertNumToAlpha = result
End Function
