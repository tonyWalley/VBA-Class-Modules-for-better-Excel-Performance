VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pLogFilePath As String
Private pTodayDate As String
Private StartTime As Double
Private pCallerName As String
' cOID
Public Property Get callerName() As String
    callerName = pCallerName
End Property

Public Property Let callerName(value As String)
    pCallerName = value
End Property

' Class Initialize to set default values for properties
Private Sub Class_Initialize()
    'pCallerName = callerName
    pTodayDate = Format(Date, "YYYYMMDD")
    
    'If [LOGS_PATH] <> "" Then
    '    pLogFilePath = [LOGS_PATH] & logType & " " & pTodayDate & ".txt"
    'Else
        pLogFilePath = getWBPath & "\Log" & " " & pTodayDate & ".txt"
    'End If
End Sub

' Method to set custom log file path if needed
Public Sub SetLogFilePath(customPath As String)
    pLogFilePath = customPath
End Sub

' General method for logging
Public Sub log(logItem As String, Optional logType As String = "Log")
    If [LoggingOn] Then
        Dim fileNo As Integer
        
        If [LOGS_PATH] <> "" Then
            pLogFilePath = [LOGS_PATH] & logType & " " & pTodayDate & ".txt"
        Else
            pLogFilePath = getWBPath & "\" & logType & " " & pTodayDate & ".txt"
        End If
        
        ' Get an available file number
        fileNo = FreeFile
        
        ' Open the log file for appending, or create it if it doesn't exist
        If dir(pLogFilePath) = "" Then
            Open pLogFilePath For Output As #fileNo
            Print #fileNo, "Log started on " & Date
        Else
            Open pLogFilePath For Append As #fileNo
        End If
        
        ' Write the log item to the file
        Print #fileNo, Now & " - " & pCallerName & " - " & logItem
        
        ' Close the file
        Close #fileNo
    End If
End Sub
' Method to log the content of a tblDic object
Public Sub logTblDic(ByRef tblDicObj As tblDic)
    Dim tblName As String
    Dim columnHeaders As String
    Dim rowData As String
    Dim logText As String
    Dim i As Long
    Dim key As Variant, colKey As Variant
    Dim lc As ListColumn
    
    With tblDicObj
        ' Retrieve the table name
        tblName = "TABLE NAME: " & .lo.name
        logText = tblName & vbCrLf
        
        ' Construct the column headers row
        For Each lc In .lo.ListColumns
            columnHeaders = columnHeaders & lc.name & vbTab
        Next lc
        logText = logText & Trim(columnHeaders) & vbCrLf
        
        ' Iterate through each row in tblDic to construct row data
        For Each key In .getTblDic.Keys
             .FindOrAddRowByKey (key)
            ' Reset rowData for each row
            rowData = ""
            For Each colKey In .rowDic.GetDictionary.Keys
                ' Use GetValue method of RowDictionaryWrapper to get cell values
                rowData = rowData & .rowDic.GetValue(CStr(colKey)) & vbTab
            Next colKey
            ' Add the constructed row data to logText
            logText = logText & Trim(rowData) & vbCrLf
        Next key
    End With
    ' Log the constructed table data
    log logText
End Sub

Public Sub logRng(rng As Range)
    Dim logText As String
    Dim row As Range
    Dim cell As Range
    Dim rowData As String
    Dim isFirstRow As Boolean: isFirstRow = True
    
    ' Check if the range is not empty
    If Not rng Is Nothing Then
        ' Iterate through each row in the range
        For Each row In rng.Rows
            ' Reset rowData for each new row
            rowData = ""
            ' Iterate through each cell in the current row
            For Each cell In row.Cells
                ' Append the cell's value to rowData, tab-delimited
                rowData = rowData & cell.value & vbTab
            Next cell
            
            ' Trim the trailing tab character from rowData
            rowData = Left(rowData, Len(rowData) - 1)
            
            ' Add the rowData to logText, ensuring not to add a newline on the first row
            If isFirstRow Then
                logText = rowData
                isFirstRow = False
            Else
                logText = logText & vbCrLf & rowData
            End If
        Next row
        
        ' Log the constructed string
        log logText
    Else
        log "Empty range passed to logRng"
    End If
End Sub

''Public Sub logTblDicInJSON(ByRef tblDicObj As tblDic)
''    Dim jsonText As String
''    Dim columnHeaders As String
''    Dim rowData As String
''    Dim i As Long
''    Dim key As Variant, colKey As Variant
''    Dim lc As ListColumn
''
''    With tblDicObj
''        ' Start JSON string
''        jsonText = "{""TableName"": """ & .LO.Name & """, "
''
''        ' Construct JSON array of column headers
''        columnHeaders = """Columns"": ["
''        For Each lc In .LO.ListColumns
''            columnHeaders = columnHeaders & """" & lc.Name & """"
''            If Not lc.Index = .LO.ListColumns.Count Then columnHeaders = columnHeaders & ", "
''        Next lc
''        columnHeaders = columnHeaders & "], "
''        jsonText = jsonText & columnHeaders
''
''        ' Construct JSON array of row data
''        jsonText = jsonText & """Rows"": ["
''        For Each key In .getTblDic.Keys
''            .FindOrAddRowByKey (key)
''            Dim currentDict As Object
''            Set currentDict = .rowDic.GetDictionary() ' Assign the dictionary to a variable
''            rowData = "{"
''            Dim lastColKey As Variant
''            If currentDict.Count > 0 Then
''                lastColKey = currentDict.Keys(currentDict.Count - 1) ' Get the last key for comparison
''            End If
''            For Each colKey In currentDict.Keys
''                ' Construct JSON object for each row
''                rowData = rowData & """" & colKey & """: """ & Replace(currentDict(colKey), """", "\""") & """"
''                If Not colKey = lastColKey Then rowData = rowData & ", "
''            Next colKey
''            rowData = rowData & "}"
''            ' Add the constructed row data to jsonText
''            jsonText = jsonText & rowData
''            If Not key = .getTblDic.Keys(.getTblDic.Count - 1) Then jsonText = jsonText & ", " ' Use variable for .getTblDic.Keys here
''        Next key
''
''        jsonText = jsonText & "]"
''
''        ' Close JSON string
''        jsonText = jsonText & "}"
''    End With
''
''    ' Log the constructed JSON string
''    log jsonText
''End Sub


' Method to start the timer
Public Sub TimerStart()
    StartTime = Timer
End Sub

' Method to end the timer and log the elapsed time
Public Sub TimerEnd()
If [PrintTimeOn] Then
    Dim EndTime As Double
    Dim ElapsedTime As Double
    Dim appCaller As Variant
    
    
    ' Capture the end time and calculate the elapsed time
    EndTime = Timer
    ElapsedTime = EndTime - StartTime
    
    ' Log the elapsed time
    log pCallerName & ": Elapsed Time: " & ElapsedTime & " seconds"
    Debug.Print pCallerName & ": Elapsed Time: " & ElapsedTime & " seconds"
End If
End Sub

' Method to handle errors
Public Sub HandleErrors()
    Dim errMsg As String
    
    log "ERROR IN " & pCallerName & " See CODE ERROR LOG"
    errMsg = "ERROR IN " & pCallerName & ": " & Err.DESCRIPTION & " (Error " & Err.Number & ")"
    
    ' Detailed error logging including system info
    Dim detailedErrorMsg As String
    detailedErrorMsg = "____________________________________________________________________________" & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "CODE ERROR in " & pCallerName & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "DESCRIPTION: " & Err.DESCRIPTION & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "ERROR NUMBER: " & Err.Number & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "WORKBOOK NAME: " & ThisWorkbook.name & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "SYS INFO: " & GetAppVersion() & " (" & Application.Version & ") " & GetExcelBitsVersion() & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "SOURCE: " & Err.source & vbCrLf
    detailedErrorMsg = detailedErrorMsg & "____________________________________________________________________________" & vbCrLf
    
    ' Log the detailed error message
    log detailedErrorMsg, "CodeError"
    
    ' Optional: Perform additional actions as needed, such as updating a cell to indicate an error
    If Not IsEmpty([warning]) Then
        [warning] = "ERROR IN " & pCallerName
    End If
End Sub

' Retrieves Excel application version as a readable string
Private Function GetAppVersion() As String
    Select Case Application.Version
        Case "16.0": GetAppVersion = "2016"
        Case "15.0": GetAppVersion = "2013"
        Case "14.0": GetAppVersion = "2010"
        Case "12.0": GetAppVersion = "2007"
        Case Else: GetAppVersion = Application.Version
    End Select
End Function

' Determines whether Excel is 32-bit or 64-bit
Private Function GetExcelBitsVersion() As String
    #If Win64 Then
        GetExcelBitsVersion = "64-bit"
    #Else
        GetExcelBitsVersion = "32-bit"
    #End If
End Function
