VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private codeTable As tblDic

Public Property Get CodeTblDic() As tblDic
    Set CodeTblDic = codeTable
End Property



' Constructor to initialize the tblDic with KEY column for direct lookups
Private Sub Class_Initialize()
    Set codeTable = New tblDic
    codeTable.AddKeyCol "KEY"  ' Initialize with KEY for direct lookups
    codeTable.buildDic b_CT.ListObjects("code_table"), valuesOnly:=True
End Sub

' Function to get a dictionary of all entries where TABLE_N matches
Public Function getCodeDic(TABLE_N As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    Dim found As Boolean
    
    For Each key In codeTable.getTblDic.Keys
        found = codeTable.FindByKey(CStr(key))
        If found And codeTable.rowDic.GetValue("TABLE_N") = TABLE_N Then
            result.Add codeTable.rowDic.GetValue("TABLE_CD"), codeTable.rowDic.GetValue("TABLE_E")
        End If
    Next key
    
    Set getCodeDic = result
End Function
' Function to get a dictionary of all entries where TABLE_N matches including property value
Public Function getFullCodeDic(TABLE_N As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    Dim found As Boolean
    
    For Each key In codeTable.getTblDic.Keys
        found = codeTable.FindByKey(CStr(key))
        If found And codeTable.rowDic.GetValue("TABLE_N") = TABLE_N Then
            result.Add codeTable.rowDic.GetValue("TABLE_CD"), getCodeDetails(TABLE_N, codeTable.rowDic.GetValue("TABLE_CD"))
        End If
    Next key
    
    Set getFullCodeDic = result
End Function
' Function to get the value of TABLE_E given TABLE_N and TABLE_CD
Public Function getTableE(TABLE_N As String, TABLE_CD As Variant) As Variant
    Dim key As String
    key = TABLE_N & TABLE_CD
    If codeTable.FindByKey(key) Then
        getTableE = codeTable.rowDic.GetValue("TABLE_E")
    Else
        getTableE = Nothing ' Return Nothing if not found
    End If
End Function

' Function to get the value of TABLE_CD given TABLE_N and TABLE_E
Public Function getTableCD(TABLE_N As String, TABLE_E As Variant) As Variant
    Dim revKey As String
    Dim key As String
    
    revKey = TABLE_N & TABLE_E
    With codeTable
        .ClearSearchCriteria ' clears any past AddRowSearch criteria
        .AddRowSearch "REV_KEY", revKey '' adds a criteria for FindRowKeyBy Search. In this case criteria= "REV_KEY item = revKey
        key = .FindRowKeyBySearch 'returns the first key found that match crireria, in the case of CodeTable, I only expect 1 match
        If key <> "" Then '' if key is found:
           .FindByKey key
            getTableCD = .rowDic.GetValue("TABLE_CD")
        Else
            getTableCD = Nothing ' Return Nothing if not found
        End If
    End With
End Function

' Function to return details as a dictionary
Public Function getCodeDetails(TABLE_N As String, TABLE_CD As Variant) As Object
    Dim details As Object
    Set details = CreateObject("Scripting.Dictionary")
    Dim key As String
    key = TABLE_N & TABLE_CD
    
    If codeTable.FindByKey(key) Then
        With codeTable.rowDic
            details.Add "TABLE_N", .GetValue("TABLE_N")
            details.Add "TABLE_CD", .GetValue("TABLE_CD")
            details.Add "TABLE_E", .GetValue("TABLE_E")
            details.Add "PROPERTY", .GetValue("PROPERTY")
        End With
    End If
    
    Set getCodeDetails = details
End Function

