VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: WorksheetDictionary

Option Explicit

Private wsDictionary As Object

Private Sub Class_Initialize()
    ' Initialize the dictionary
    Set wsDictionary = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        wsDictionary.Add ws.codeName, ws
    Next ws
End Sub


Private Sub Class_Terminate()
    ' Clean up
    Set wsDictionary = Nothing
End Sub

' Method to get a worksheet by its CodeName
Public Function GetWorksheetByCodeName(codeName As String) As Worksheet
    If wsDictionary.Exists(codeName) Then
        Set GetWorksheetByCodeName = wsDictionary(codeName)
    Else
        Set GetWorksheetByCodeName = Nothing
    End If
End Function


' New method to get worksheets by type based on the first letter of the CodeName
Public Function wsDicByType(WSType As String) As Scripting.Dictionary
    Dim resultDict As Scripting.Dictionary
    Dim key As Variant
    Set resultDict = New Scripting.Dictionary
    
    Dim ws As Worksheet
    Dim sheetCode As String
    
    ' Loop through all worksheets in the dictionary
    For Each key In wsDictionary.Keys
        Set ws = wsDictionary(key)
        ' Check if the CodeName starts with WSType
        If Left(ws.codeName, 1) = WSType Then
            ' Extract sheetCode
            sheetCode = Split(ws.codeName, "_")(1)
            ' Add to result dictionary if not already present
            If Not resultDict.Exists(sheetCode) Then
                resultDict.Add sheetCode, ws
            End If
        End If
    Next key
    
    ' Return the filtered dictionary
    Set wsDicByType = resultDict
End Function


