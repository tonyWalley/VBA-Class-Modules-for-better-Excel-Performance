VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxTbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: ListBoxTbl
Option Explicit

Private codeTable As codeTable
Private strMgr As StringMgr
Private listBoxName As String
Private internalItemStr As String
Private valuesDic As Scripting.Dictionary
Private propsDic As Scripting.Dictionary

Public Enum strPadType
    padLeft = 1
    padRight = 2
End Enum

' Class Initialization
Private Sub Class_Initialize()
    Set codeTable = New codeTable
    Set strMgr = New StringMgr
    Set valuesDic = New Scripting.Dictionary
    Set propsDic = New Scripting.Dictionary
End Sub

' Property for ListBox name
Public Property Get ListBoxNm() As String
    ListBoxNm = listBoxName
End Property

Public Property Let ListBoxNm(value As String)
    listBoxName = value
    Set propsDic = codeTable.getFullCodeDic(listBoxName)
    BuildBlankItemStr
    InitValuesDic
End Property

' Property for the Item String of the ListBox
Public Property Get itemStr() As String
    itemStr = internalItemStr
End Property

Public Property Let itemStr(value As String)
    internalItemStr = value
    ParseItemStr
End Property

' Method to set the value of a column in the list box
Public Sub SetColumnValue(ColumnName As String, value As Variant)
    Dim columnInfo As Scripting.Dictionary
    Dim orderAndLen As Variant
    
    Set columnInfo = propsDic(ColumnName)
    orderAndLen = Split(columnInfo("TABLE_E"), ".")
    
    ' Storing the raw value in valuesDic
    valuesDic(ColumnName) = value
    
    ' Update ItemStr based on the updated ValuesDic
    UpdateItemStr
End Sub

' Method to rebuild the full item string based on the current values
Private Sub UpdateItemStr()
    Dim key As Variant
    Dim columnInfo As Scripting.Dictionary
    Dim orderAndLen As Variant
    Dim colLen As Integer
    Dim padType As strPadType
    Dim paddedValue As String
    Dim strValue As String
    
    internalItemStr = ""
    For Each key In propsDic.Keys
        Set columnInfo = propsDic(key)
        orderAndLen = Split(columnInfo("TABLE_E"), ".")
        colLen = CInt(orderAndLen(1))
        
        Select Case columnInfo("PROPERTY")
            Case "txt"
                strValue = valuesDic(key)
                padType = strPadType.padRight
            Case "num", "dec"
                strValue = CStr(valuesDic(key))
                padType = strPadType.padLeft
            Case "yesno"
                strValue = IIf(CBool(valuesDic(key)), "Y", "N")
                padType = strPadType.padRight
        End Select

        paddedValue = strMgr.PadAndTrim(strValue, padType, colLen)
        internalItemStr = internalItemStr & paddedValue & "|"
    Next key
    
    ' Remove the trailing '|'
    If Right(internalItemStr, 1) = "|" Then
        internalItemStr = Left(internalItemStr, Len(internalItemStr) - 1)
    End If
End Sub


' Method to parse the item string into values based on column properties
Private Sub ParseItemStr()
    Dim parts As Variant
    Dim i As Integer
    Dim columnInfo As Scripting.Dictionary
    Dim value As Variant
    Dim yesnoVal As Boolean

    parts = strMgr.parseAndTrim(internalItemStr, "|")

    For i = 0 To UBound(parts)
        Set columnInfo = propsDic(propsDic.Keys(i))
        Select Case columnInfo("PROPERTY")
            Case "txt"
                value = Trim(parts(i))
            Case "num"
                value = CLng(Trim(parts(i)))
            Case "dec"
                value = CDbl(Trim(parts(i)))
            Case "yesno"
                yesnoVal = (Trim(parts(i)) = "Y")
                value = yesnoVal
        End Select
        valuesDic(propsDic.Keys(i)) = value
    Next i
End Sub


' Method to build a blank item string based on the default lengths and order from propsDic
Private Sub BuildBlankItemStr()
    Dim key As Variant
    Dim orderAndLen As Variant
    Dim colLen As Integer
    Dim colProps As Scripting.Dictionary

    
    internalItemStr = ""

    ' Ensure propsDic is ordered or has a predictable iteration if order matters
    For Each key In propsDic
        Set colProps = propsDic(key)
        orderAndLen = Split(colProps("TABLE_E"), ".")

        'orderAndLen = Split(propsDic(key)("TABLE_E"), ".")
        colLen = CInt(orderAndLen(1))  ' Get the length of the column from ORDER.LEN
        internalItemStr = internalItemStr & String(colLen, " ") & "|"
    Next key
    
    ' Remove the trailing '|' to match exactly the desired format
    If Right(internalItemStr, 1) = "|" Then
        internalItemStr = Left(internalItemStr, Len(internalItemStr) - 1)
    End If
End Sub
Private Sub InitValuesDic()
Dim key As Variant

    For Each key In propsDic
        valuesDic.Add CStr(key), Empty
    Next key
End Sub

' Method to retrieve the dictionary of values
Public Function GetValueDic() As Scripting.Dictionary
    Set GetValueDic = valuesDic
End Function



