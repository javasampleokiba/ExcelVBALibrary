Attribute VB_Name = "CellAddressUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : セルアドレスに関するユーティリティモジュール
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : 列名を列番号に変換します。
'            実際のExcelの最大列数を超えるような列名が指定された場合も変換します。
'            変換できない列名が指定された場合は、0を返します。
'
' PARAMS   : columnName - 列名
'
' RETURN   : 列番号
'
'------------------------------------------------------------------------------
Public Function ToColumnIndex(ByVal columnName As String) As Long
    Dim i           As Long
    Dim result      As Long
    Dim code        As Integer

    columnName = UCase(columnName)

    For i = 1 To Len(columnName)
        code = Asc(mid(columnName, i, 1))

        ' A～Z以外の場合
        If code < 65 Or 90 < code Then
            ToColumnIndex = 0
            Exit Function
        End If

        result = result + (code - 64) * (26 ^ (Len(columnName) - i))
    Next

    ToColumnIndex = result

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 列番号を列名に変換します。
'            実際のExcelの最大列数を超えるような列番号が指定された場合も変換します。
'            変換できない番号が指定された場合は、空文字を返します。
'
' PARAMS   : column - 列番号
'
' RETURN   : 列名
'
'------------------------------------------------------------------------------
Public Function ToColumnName(ByVal column As Long) As String
    Dim result      As String

    If column < 1 Then
        ToColumnName = ""
        Exit Function
    End If

    Do
        column = column - 1
        result = Chr(65 + (column Mod 26)) & result
        column = column \ 26

    Loop Until column = 0

    ToColumnName = result

End Function
