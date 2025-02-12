VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowDictionaryWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: RowDictionaryWrapper
Option Explicit

Private ColData As Object ' Dictionary to store column data or formulas
Private UpdatedCols As Collection
Private pIsLastItem As Boolean

Public Property Get dic() As Scripting.Dictionary
    Set dic = ColData
End Property

' Property to get/set the isLastItem flag
Public Property Get isLastItem() As Boolean
    isLastItem = pIsLastItem
End Property

Public Property Let isLastItem(value As Boolean)
    pIsLastItem = value
End Property
' Initialize the class
Private Sub Class_Initialize()
    Set ColData = CreateObject("Scripting.Dictionary")
    Set UpdatedCols = New Collection
     pIsLastItem = False
End Sub

' Initialize the wrapper with column formulas
Public Sub InitializeWithFormulas(columnFormulas As Object)
    Dim ColName As Variant
    For Each ColName In columnFormulas.Keys
        ' Store the formula or default value for each column
        ColData(ColName) = columnFormulas(ColName)
    Next ColName
End Sub
Public Sub InitializeFromExisting(lr As listRow, valuesOnly As Boolean)
    Dim col As ListColumn
    Dim cell As Range
    ' Clear any existing data in colData
    ColData.RemoveAll
    
    ' Iterate through each column in the ListRow
    For Each col In lr.Parent.ListColumns
        ' Get the cell in the ListRow corresponding to the current column
        Set cell = lr.Range.Cells(1, col.Index)
            If valuesOnly Then
                ColData(col.name) = cell.value
            Else
                ColData(col.name) = cell.Formula
            End If
    Next col
End Sub


' Set value for a specific column
Public Sub SetValue(ColumnName As String, value As Variant)
    ' Ensure the column name is in the dictionary before setting its value
    If Not ColData.Exists(ColumnName) Then
        ColData.Add ColumnName, value
        UpdatedCols.Add ColumnName, ColumnName
    Else
        If ColData(ColumnName) <> value Then ' Check if the value is actually being updated
            ColData(ColumnName) = value
            On Error Resume Next ' Prevent error on adding a duplicate
            UpdatedCols.Add ColumnName, ColumnName
            On Error GoTo 0 ' Turn back on error handling
        End If
    End If
End Sub
Public Property Get UpdatedColumns() As Collection
    Set UpdatedColumns = UpdatedCols
End Property

' Get value for a specific column
Public Function GetValue(ColumnName As String) As Variant
    If ColData.Exists(ColumnName) Then
        GetValue = ColData(ColumnName)
    Else
        GetValue = vbNullString ' Or any other appropriate default value
    End If
End Function

' Check if a specific column exists
Public Function Exists(ColumnName As String) As Boolean
    Exists = ColData.Exists(ColumnName)
End Function

' Expose the underlying dictionary for direct operations (e.g., iterating keys)
Public Function GetDictionary() As Object
    Set GetDictionary = ColData
End Function

Public Function ToArray(lstObj As ListObject) As Variant
    Dim result() As Variant
    ReDim result(1 To 1, 1 To lstObj.ListColumns.Count)
    
    Dim colIndex As Integer
    For colIndex = 1 To lstObj.ListColumns.Count
        Dim ColName As String
        ColName = lstObj.ListColumns(colIndex).name
        If Me.GetDictionary().Exists(ColName) Then
            result(1, colIndex) = Me.GetDictionary()(ColName)
        Else
            result(1, colIndex) = "" ' Handle missing values or use a default
        End If
    Next colIndex
    
    ToArray = result
End Function
