Attribute VB_Name = "CellAddressUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : �Z���A�h���X�Ɋւ��郆�[�e�B���e�B���W���[��
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : �񖼂��ԍ��ɕϊ����܂��B
'            ���ۂ�Excel�̍ő�񐔂𒴂���悤�ȗ񖼂��w�肳�ꂽ�ꍇ���ϊ����܂��B
'            �ϊ��ł��Ȃ��񖼂��w�肳�ꂽ�ꍇ�́A0��Ԃ��܂��B
'
' PARAMS   : columnName - ��
'
' RETURN   : ��ԍ�
'
'------------------------------------------------------------------------------
Public Function ToColumnIndex(ByVal columnName As String) As Long
    Dim i           As Long
    Dim result      As Long
    Dim code        As Integer

    columnName = UCase(columnName)

    For i = 1 To Len(columnName)
        code = Asc(mid(columnName, i, 1))

        ' A�`Z�ȊO�̏ꍇ
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
' FUNCTION : ��ԍ���񖼂ɕϊ����܂��B
'            ���ۂ�Excel�̍ő�񐔂𒴂���悤�ȗ�ԍ����w�肳�ꂽ�ꍇ���ϊ����܂��B
'            �ϊ��ł��Ȃ��ԍ����w�肳�ꂽ�ꍇ�́A�󕶎���Ԃ��܂��B
'
' PARAMS   : column - ��ԍ�
'
' RETURN   : ��
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
