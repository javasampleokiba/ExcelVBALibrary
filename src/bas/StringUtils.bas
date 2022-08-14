Attribute VB_Name = "StringUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : �����񑀍�Ɋւ��郆�[�e�B���e�B���W���[��
'
' NOTE   : �T���Q�[�g�y�A�ɂ͑Ή����Ă��܂���B
'
' LINK   : ���̃��W���[���͈ȉ��̃��W���[�����Q�Ƃ��Ă��܂��B
'
'          �EArrayUtils
'------------------------------------------------------------------------------

' �󔒕���
Private Const CHARS_BLANK As String = " �@" & vbTab & vbCr & vbLf
' ����
Private Const CHARS_DIGIT As String = "0123456789"
' �p��
Private Const CHARS_ALPHA As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������ŏI���Ȃ��ꍇ�A�����ɕt�����ĕԂ��܂��B
'
' PARAMS   : str          - �Ώە�����
'            suffix       - �ǉ����镶����
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�������t������������
'------------------------------------------------------------------------------
Public Function AppendIfMissing(ByVal str As String, ByVal suffix As String, _
                                Optional ByVal ignoreCase As Boolean = False) As String

    If EndsWith(str, suffix, ignoreCase) Then
        AppendIfMissing = str
    Else
        AppendIfMissing = str & suffix
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������Ɏw�肳�ꂽ�����񂪊܂܂�邩���肵�܂��B
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�����񂪊܂܂��ꍇ�� True
'------------------------------------------------------------------------------
Public Function Contains(ByVal str As String, ByVal searchStr As String, _
                        Optional ByVal ignoreCase As Boolean = False) As Boolean

    Contains = 0 < IndexOf(str, searchStr, 1, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������Ɏw�肳�ꂽ�����񂷂ׂĂ��܂܂�邩���肵�܂��B
'
' PARAMS   : str        - �Ώە�����
'            searchStrs - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ�����񂷂ׂĂ��܂܂��ꍇ�� True
'------------------------------------------------------------------------------
Public Function ContainsAll(ByVal str As String, ParamArray searchStrs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, searchStrs)

    ContainsAll = False

    If ArrayUtils.IsEmptyArray(params) Then
        Exit Function
    End If

    For Each param In params
        If Not Contains(str, param) Then
            Exit Function
        End If
    Next

    ContainsAll = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������Ɏw�肳�ꂽ������̂����ꂩ���܂܂�邩���肵�܂��B
'
' PARAMS   : str        - �Ώە�����
'            searchStrs - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ������̂����ꂩ���܂܂��ꍇ�� True
'------------------------------------------------------------------------------
Public Function ContainsAny(ByVal str As String, ParamArray searchStrs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, searchStrs)

    ContainsAny = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If Contains(str, param) Then
                Exit Function
            End If
        Next
    End If

    ContainsAny = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������Ɏw�肳�ꂽ�����񂪂����܂܂�邩�J�E���g���܂��B
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�����񂪌���������
'------------------------------------------------------------------------------
Public Function Count(ByVal str As String, ByVal searchStr As String, _
                        Optional ByVal ignoreCase As Boolean = False) As Long
    Dim idx     As Long
    Dim cnt     As Long

    idx = 1
    cnt = 0

    Do While True
        idx = IndexOf(str, searchStr, idx, ignoreCase)
        If idx = 0 Then
            Exit Do
        End If
        idx = idx + Len(searchStr)
        cnt = cnt + 1
    Loop

    Count = cnt

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������ŏI��邩�ǂ����𔻒肵�܂��B
'
' PARAMS   : str          - �Ώە�����
'            suffix       - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ������ŏI���ꍇ�� True
'------------------------------------------------------------------------------
Public Function EndsWith(ByVal str As String, ByVal suffix As String, _
                        Optional ByVal ignoreCase As Boolean = False) As Boolean

    If str = "" Or suffix = "" Or Len(str) < Len(suffix) Then
        EndsWith = False
    Else
        EndsWith = LastIndexOf(str, suffix, -1, ignoreCase) = Len(str) - Len(suffix) + 1
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������̂����ꂩ�ŏI��邩�ǂ����𔻒肵�܂��B
'
' PARAMS   : str      - �Ώە�����
'            suffixes - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ������ŏI���ꍇ�� True
'------------------------------------------------------------------------------
Public Function EndsWithAny(ByVal str As String, ParamArray suffixes() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, suffixes)

    EndsWithAny = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If EndsWith(str, param) Then
                Exit Function
            End If
        Next
    End If

    EndsWithAny = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ2�̕����񂪓��������𔻒肵�܂��B
'
' PARAMS   : str1         - ����Ώە�����1
'            str2         - ����Ώە�����2
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : 2�̕����񂪓������ꍇ�� True
'------------------------------------------------------------------------------
Public Function Equals(ByVal str1 As String, ByVal str2 As String, _
                        Optional ByVal ignoreCase As Boolean = False) As Boolean

    If ignoreCase Then
        str1 = LCase(str1)
        str2 = LCase(str2)
    End If

    Equals = str1 = str2

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������̂����ꂩ�Ɠ��������𔻒肵�܂��B
'
' PARAMS   : str  - ����Ώە�����
'            strs - ����Ώە�����ꗗ
'
' RETURN   : �����ꂩ�̕����񂪓������ꍇ�� True
'------------------------------------------------------------------------------
Public Function EqualsAny(ByVal str As String, ParamArray strs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    EqualsAny = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If Equals(str, param) Then
                Exit Function
            End If
        Next
    End If

    EqualsAny = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ������̂����A�󔒕�����ł͂Ȃ��ŏ��̕������Ԃ��܂��B
'            (������Ȃ��ꍇ�͋󕶎���Ԃ��܂�)
'
' PARAMS   : strs - ��╶����ꗗ
'
' RETURN   : �󔒕�����ł͂Ȃ�������
'------------------------------------------------------------------------------
Public Function FirstNotBlank(ParamArray strs() As Variant) As String
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    FirstNotBlank = ""

    If ArrayUtils.IsEmptyArray(params) Then
        Exit Function
    End If

    For Each param In params
        If Not IsBlank(param) Then
            FirstNotBlank = param
            Exit Function
        End If
    Next

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������Ŏw�肳�ꂽ�����񂪍ŏ��Ɍ����ʒu�C���f�b�N�X��Ԃ��܂��B
'            (�擪�̏ꍇ��1�ł��B������Ȃ��ꍇ��0��Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [start]      - �����J�n�ʒu (�ȗ�����1)
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�����񂪍ŏ��Ɍ����ʒu�C���f�b�N�X
'------------------------------------------------------------------------------
Public Function IndexOf(ByVal str As String, ByVal searchStr As String, _
                        Optional ByVal start As Long = 1, _
                        Optional ByVal ignoreCase As Boolean = False) As Long

    If str = "" Or searchStr = "" Then
        IndexOf = 0
        Exit Function
    End If

    If ignoreCase Then
        str = LCase(str)
        searchStr = LCase(searchStr)
    End If

    IndexOf = InStr(start, str, searchStr)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������Ŏw�肳�ꂽ������̂����ꂩ���ŏ��Ɍ����
'            �ʒu�C���f�b�N�X��Ԃ��܂��B
'            (�擪�̏ꍇ��1�ł��B������Ȃ��ꍇ��0��Ԃ��܂�)
'
' PARAMS   : str        - �Ώە�����
'            searchStrs - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ������̂����ꂩ���ŏ��Ɍ����ʒu�C���f�b�N�X
'------------------------------------------------------------------------------
Public Function IndexOfAny(ByVal str As String, ParamArray searchStrs() As Variant) As Long
    Dim params()    As Variant
    Dim param       As Variant
    Dim idx         As Long
    Dim minIdx      As Long

    params = ArrayUtils.Flatten(0, searchStrs)

    minIdx = 0

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            idx = IndexOf(str, param)
            If 0 < idx And (minIdx = 0 Or idx < minIdx) Then
                minIdx = idx
            End If
        Next
    End If

    IndexOfAny = minIdx

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���ׂĂ̕����񂪋󔒕����݂̂ō\������Ă��邩���肵�܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : strs - �Ώە�����ꗗ
'
' RETURN   : ���ׂĂ��󔒕����݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAllBlank(ParamArray strs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    IsAllBlank = False

    If ArrayUtils.IsEmptyArray(params) Then
        Exit Function
    End If

    For Each param In params
        If Not IsBlank(param) Then
            Exit Function
        End If
    Next

    IsAllBlank = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���ׂĂ̕����񂪋󕶎������肵�܂��B
'
' PARAMS   : strs - �Ώە�����ꗗ
'
' RETURN   : ���ׂĂ̕����񂪋󕶎��̏ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAllEmpty(ParamArray strs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    IsAllEmpty = False

    If ArrayUtils.IsEmptyArray(params) Then
        Exit Function
    End If

    For Each param In params
        If param <> "" Then
            Exit Function
        End If
    Next

    IsAllEmpty = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪉p���݂̂ō\������Ă��邩���肵�܂��B
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �p���݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAlpha(ByVal str As String) As Boolean
    Dim i       As Long
    Dim c       As String

    IsAlpha = False

    If str = "" Then
        Exit Function
    End If

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If InStr(CHARS_ALPHA, c) = 0 Then
            Exit Function
        End If
    Next

    IsAlpha = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪉p�����݂̂ō\������Ă��邩���肵�܂��B
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �p�����݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAlphaDigit(ByVal str As String) As Boolean
    Dim i       As Long
    Dim c       As String

    IsAlphaDigit = False

    If str = "" Then
        Exit Function
    End If

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If InStr(CHARS_ALPHA, c) = 0 And InStr(CHARS_DIGIT, c) = 0 Then
            Exit Function
        End If
    Next

    IsAlphaDigit = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����ꂩ�̕����񂪋󔒕����݂̂ō\������Ă��邩���肵�܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : strs - �Ώە�����ꗗ
'
' RETURN   : �����ꂩ���󔒕����݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAnyBlank(ParamArray strs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    IsAnyBlank = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If IsBlank(param) Then
                Exit Function
            End If
        Next
    End If

    IsAnyBlank = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����ꂩ�̕����񂪋󕶎������肵�܂��B
'
' PARAMS   : strs - �Ώە�����ꗗ
'
' RETURN   : �����ꂩ�̕����񂪋󕶎��̏ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsAnyEmpty(ParamArray strs() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, strs)

    IsAnyEmpty = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If param = "" Then
                Exit Function
            End If
        Next
    End If

    IsAnyEmpty = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪋󔒕����݂̂ō\������Ă��邩���肵�܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �󔒕����݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsBlank(ByVal str As String) As Boolean
    Dim i       As Long
    Dim c       As String

    IsBlank = False

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If InStr(CHARS_BLANK, c) = 0 Then
            Exit Function
        End If
    Next

    IsBlank = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪐����݂̂ō\������Ă��邩���肵�܂��B
'            ("0"�`"9"�𐔎��Ƃ݂Ȃ��܂�)
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �����݂̂ō\������Ă���ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsDigit(ByVal str As String) As Boolean
    Dim i       As Long
    Dim c       As String

    IsDigit = False

    If str = "" Then
        Exit Function
    End If

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If InStr(CHARS_DIGIT, c) = 0 Then
            Exit Function
        End If
    Next

    IsDigit = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������Ŏw�肳�ꂽ�����񂪍Ō�Ɍ����ʒu�C���f�b�N�X��Ԃ��܂��B
'            (������Ȃ��ꍇ��0��Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [start]      - �����J�n�ʒu (�ȗ����͖���)
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�����񂪍Ō�Ɍ����ʒu�C���f�b�N�X
'------------------------------------------------------------------------------
Public Function LastIndexOf(ByVal str As String, ByVal searchStr As String, _
                            Optional ByVal start As Long = -1, _
                            Optional ByVal ignoreCase As Boolean = False) As Long

    If str = "" Or searchStr = "" Then
        LastIndexOf = 0
        Exit Function
    End If

    If ignoreCase Then
        str = LCase(str)
        searchStr = LCase(searchStr)
    End If

    LastIndexOf = InStrRev(str, searchStr, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������Ŏw�肳�ꂽ������̂����ꂩ���Ō�Ɍ����
'            �ʒu�C���f�b�N�X��Ԃ��܂��B
'            (������Ȃ��ꍇ��0��Ԃ��܂�)
'
' PARAMS   : str        - �Ώە�����
'            searchStrs - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ������̂����ꂩ���Ō�Ɍ����ʒu�C���f�b�N�X
'------------------------------------------------------------------------------
Public Function LastIndexOfAny(ByVal str As String, ParamArray searchStrs() As Variant) As Long
    Dim params()    As Variant
    Dim param       As Variant
    Dim idx         As Long
    Dim maxIdx      As Long

    params = ArrayUtils.Flatten(0, searchStrs)

    maxIdx = 0

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            idx = LastIndexOf(str, param)
            If maxIdx < idx Then
                maxIdx = idx
            End If
        Next
    End If

    LastIndexOfAny = maxIdx

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂩂�w�肳�ꂽ��������O�̕����������Ԃ��܂��B
'            (������Ȃ��ꍇ�͋󕶎���Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ��������O�̕���������
'------------------------------------------------------------------------------
Public Function LeftBefore(ByVal str As String, ByVal searchStr As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String
    Dim idx     As Long

    idx = IndexOf(str, searchStr, 1, ignoreCase)

    If 0 < idx Then
        LeftBefore = left(str, idx - 1)
    Else
        LeftBefore = ""
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ�����񂩂�擪�̋󔒕�������폜���ĕԂ��܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �擪�̋󔒕�������폜����������
'------------------------------------------------------------------------------
Public Function LTrim(ByVal str As String) As String
    Dim i       As Long
    Dim c       As String

    If str = "" Then
        LTrim = str
        Exit Function
    End If

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If Not IsBlank(c) Then
            Exit For
        End If
    Next

    LTrim = right(str, Len(str) - i + 1)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂩂�w�肳�ꂽ�J�n�E�I��������̊Ԃ̕����������Ԃ��܂��B
'            (������Ȃ��ꍇ�͋󕶎���Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            beforeStr    - �J�n������
'            afterStr     - �I��������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�J�n�E�I��������̊Ԃ̕���������
'------------------------------------------------------------------------------
Public Function MidBetween(ByVal str As String, ByVal beforeStr As String, _
                            ByVal afterStr As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String

    MidBetween = LeftBefore(RightAfter(str, beforeStr, ignoreCase), afterStr, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������̎w��ʒu�̕�������A�w�肵���ʂ̕�����ɒu�������܂��B
'            (�w�肵���͈͂��͈͊O�̏ꍇ�͉����������̂܂ܕԂ��܂�)
'
' PARAMS   : str        - �Ώە�����
'            replaceStr - �u�����镶����
'            startIdx   - �����J�n�ʒu
'            endIdx     - �����I���ʒu
'
' RETURN   : �u����̕�����
'------------------------------------------------------------------------------
Public Function Overlay(ByVal str As String, ByVal replaceStr As String, _
                        ByVal startIdx As Long, ByVal endIdx As Long) As String
    Dim buf     As String

    If endIdx < startIdx Or Len(str) < startIdx Or endIdx < 1 Then
        Overlay = str
        Exit Function
    End If

    If startIdx < 1 Then
        buf = ""
    Else
        buf = left(str, startIdx - 1)
    End If

    If endIdx <= Len(str) Then
        Overlay = buf & replaceStr & right(str, Len(str) - endIdx)
    Else
        Overlay = buf & replaceStr
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ���������؂蕶���񂪍ŏ��Ɍ����ʒu��3�ɕ������A
'            ��؂���O�A��؂蕶����A��؂������i�[�����z���Ԃ��܂��B
'            ��؂蕶���񂪌�����Ȃ��ꍇ�́A2�`3�Ԗڂ̗v�f���󕶎��ɂȂ�܂��B
'
' PARAMS   : str          - �Ώە�����
'            separator    - ��؂蕶����
'            [start]      - �����J�n�ʒu (�ȗ�����1)
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : ��؂���O�A��؂蕶����A��؂������i�[�����z��
'------------------------------------------------------------------------------
Public Function Partition(ByVal str As String, ByVal separator As String, _
                            Optional ByVal start As Long = 1, _
                            Optional ByVal ignoreCase As Boolean = False) As String()
    Dim result(0 To 2)  As String
    Dim idx             As Long

    If separator = "" Then
        result(0) = ""
        result(1) = ""
        result(2) = str
    Else
        idx = IndexOf(str, separator, start, ignoreCase)
        If idx = 0 Then
            result(0) = str
            result(1) = ""
            result(2) = ""
        Else
            result(0) = left(str, idx - 1)
            result(1) = Mid(str, idx, Len(separator))
            result(2) = right(str, Len(str) - idx - Len(separator) + 1)
        End If
    End If

    Partition = result

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������Ŏn�܂�Ȃ��ꍇ�A�擪�ɕt�����ĕԂ��܂��B
'
' PARAMS   : str          - �Ώە�����
'            preffix      - �ǉ����镶����
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�������t������������
'------------------------------------------------------------------------------
Public Function PrependIfMissing(ByVal str As String, ByVal preffix As String, _
                                Optional ByVal ignoreCase As Boolean = False) As String

    If StartsWith(str, preffix, ignoreCase) Then
        PrependIfMissing = str
    Else
        PrependIfMissing = preffix & str
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂩂�w�肵����������폜���܂��B
'
' PARAMS   : str          - �Ώە�����
'            removeStr    - �폜���镶����
'            [start]      - �����J�n�ʒu (�ȗ�����1)
'            [size]       - �폜����ő吔 (�ȗ��A�܂���0���w�肷��Ƃ��ׂč폜)
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �폜��̕�����
'------------------------------------------------------------------------------
Public Function Remove(ByVal str As String, ByVal removeStr As String, _
                        Optional ByVal start As Long = 1, Optional ByVal size As Long = 0, _
                        Optional ByVal ignoreCase As Boolean = False) As String

    Remove = Replace(str, removeStr, "", start, size, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������ŏI���ꍇ�A���̕�����𖖔������菜���܂��B
'            (������Ȃ��ꍇ�͉��������Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            suffix       - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ������𖖔�����폜����������
'------------------------------------------------------------------------------
Public Function RemoveEnd(ByVal str As String, ByVal suffix As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String

    If EndsWith(str, suffix, ignoreCase) Then
        RemoveEnd = left(str, Len(str) - Len(suffix))
    Else
        RemoveEnd = str
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������Ŏn�܂�ꍇ�A���̕������擪�����菜���܂��B
'            (������Ȃ��ꍇ�͉��������Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            prefix       - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ�������擪����폜����������
'------------------------------------------------------------------------------
Public Function RemoveStart(ByVal str As String, ByVal prefix As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String

    If StartsWith(str, prefix, ignoreCase) Then
        RemoveStart = right(str, Len(str) - Len(prefix))
    Else
        RemoveStart = str
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������̐擪����w�肵����������������A�w�肵���ʂ̕�����ɒu�������܂��B
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            replaceStr   - �u�����镶����
'            [start]      - �����J�n�ʒu (�ȗ�����1)
'            [size]       - �u������ő吔 (�ȗ��A�܂���0���w�肷��Ƃ��ׂĒu��)
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �u����̕�����
'------------------------------------------------------------------------------
Public Function Replace(ByVal str As String, ByVal searchStr As String, _
                        ByVal replaceStr As String, Optional ByVal start As Long = 1, _
                        Optional ByVal size As Long = 0, _
                        Optional ByVal ignoreCase As Boolean = False) As String
    Dim cnt     As Long
    Dim parts() As String
    Dim rest    As String
    Dim buf     As String

    If str = "" Then
        Replace = ""
        Exit Function
    End If

    If searchStr = "" Then
        Replace = str
        Exit Function
    End If

    cnt = 0
    rest = str

    Do While True
        parts = Partition(rest, searchStr, start, ignoreCase)
        cnt = cnt + 1
        start = 1
        rest = parts(2)
        If rest = "" Then
            buf = buf & parts(0)
            Exit Do
        ElseIf cnt = size Then
            buf = buf & parts(0) & replaceStr & parts(2)
            Exit Do
        Else
            buf = buf & parts(0) & replaceStr
        End If
    Loop

    Replace = buf

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ������̕��т𔽓]���܂��B
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : ���]��̕�����
'------------------------------------------------------------------------------
Public Function Reverse(ByVal str As String) As String
    Dim i       As Long
    Dim c       As String
    Dim buf     As String

    For i = Len(str) To 1 Step -1
        c = Mid(str, i, 1)
        buf = buf & c
    Next

    Reverse = buf

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂩂�w�肳�ꂽ���������̕����������Ԃ��܂��B
'            (������Ȃ��ꍇ�͋󕶎���Ԃ��܂�)
'
' PARAMS   : str          - �Ώە�����
'            searchStr    - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ���������̕���������
'------------------------------------------------------------------------------
Public Function RightAfter(ByVal str As String, ByVal searchStr As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String
    Dim idx     As Long

    idx = IndexOf(str, searchStr, 1, ignoreCase)

    If 0 < idx Then
        RightAfter = right(str, Len(str) - idx - Len(searchStr) + 1)
    Else
        RightAfter = ""
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������w�肳�ꂽ������������](�ړ�)�����܂��B
'            �����ɐ��̒l���w�肷��ƌ������Ɉړ����A
'            ���̒l���w�肷��ƑO�����Ɉړ����܂��B
'
' PARAMS   : str      - �Ώە�����
'            distance - ������(�ړ�����)
'
' RETURN   : ��]��̕�����
'------------------------------------------------------------------------------
Public Function Rotate(ByVal str As String, ByVal distance As Long) As String
    Dim l       As Long
    Dim dst     As Long
    Dim buf     As String

    If str = "" Then
        Rotate = str
        Exit Function
    End If

    l = Len(str)
    dst = distance Mod l

    If 0 <= dst Then
        Rotate = right(str, dst) & left(str, l - dst)
    Else
        dst = -dst
        Rotate = right(str, l - dst) & left(str, dst)
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ�����񂩂疖���̋󔒕�������폜���ĕԂ��܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �����̋󔒕�������폜����������
'------------------------------------------------------------------------------
Public Function RTrim(ByVal str As String) As String
    Dim i       As Long
    Dim c       As String

    If str = "" Then
        RTrim = str
        Exit Function
    End If

    For i = Len(str) To 1 Step -1
        c = Mid(str, i, 1)
        If Not IsBlank(c) Then
            Exit For
        End If
    Next

    RTrim = left(str, i)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������p�����������؂�Ƃ��ĕ������A���ʂ�z��Ɋi�[���ĕԂ��܂��B
'            ��؂蕶�����A�����Ă������́A�ЂƂ̋�؂蕶����Ƃ��Ĉ����܂��B
'            ��F"1abc2def3"�̏ꍇ�́A(1, 2, 3)���Ԃ����B
'
' PARAMS   : str                 - �Ώە�����
'            [containsSeparator] - ��؂蕶��������ʂɊ܂߂邩 (�ȗ����� False)
'
' RETURN   : �����������ʂ��i�[�����z��
'------------------------------------------------------------------------------
Public Function SplitByAlpha(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByAlpha = SplitByChars(str, CHARS_ALPHA, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������󔒕��������؂�Ƃ��ĕ������A���ʂ�z��Ɋi�[���ĕԂ��܂��B
'            ��؂蕶�����A�����Ă������́A�ЂƂ̋�؂蕶����Ƃ��Ĉ����܂��B
'            ��F"a   b�@�@c"�̏ꍇ�́A(a, b, c)���Ԃ����B
'
' PARAMS   : str                 - �Ώە�����
'            [containsSeparator] - ��؂蕶��������ʂɊ܂߂邩 (�ȗ����� False)
'
' RETURN   : �����������ʂ��i�[�����z��
'------------------------------------------------------------------------------
Public Function SplitByBlank(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByBlank = SplitByChars(str, CHARS_BLANK, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ��������w�肵�������ꂩ�̕�������؂�Ƃ��ĕ������A
'            ���ʂ�z��Ɋi�[���ĕԂ��܂��B
'            ��؂蕶�����A�����Ă������́A�ЂƂ̋�؂蕶����Ƃ��Ĉ����܂��B
'            ��F"abc-|-|-def"�ɑ΂��āA��؂蕶���ꗗ��"-|"���w�肵���ꍇ�A
'                (abc, def)���Ԃ����B
'
' PARAMS   : str                 - �Ώە�����
'            separateChars       - ��؂蕶���ꗗ
'            [containsSeparator] - ��؂蕶��������ʂɊ܂߂邩 (�ȗ����� False)
'
' RETURN   : �����������ʂ��i�[�����z��
'------------------------------------------------------------------------------
Public Function SplitByChars(ByVal str As String, ByVal separateChars As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()
    Dim i           As Long
    Dim st          As Long
    Dim en          As Long
    Dim c           As String
    Dim subStr      As String
    Dim result()    As String

    If str = "" Then
        Call ArrayUtils.Push(result, "")
        SplitByChars = result
        Exit Function
    End If

    st = 0
    en = 0

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        ' ��؂蕶�����������Ă��Ȃ��ꍇ
        If st = 0 Then
            ' ��؂蕶���̏ꍇ
            If 0 < InStr(separateChars, c) Then
                st = i
            End If
        ' ��؂蕶�����������Ă���ꍇ
        Else
            ' ��؂蕶���ł͂Ȃ��ꍇ
            If InStr(separateChars, c) = 0 Then
                subStr = Mid(str, en + 1, st - en - 1)
                If subStr <> "" Then
                    Call ArrayUtils.Push(result, subStr)
                End If
                en = i - 1
                If containsSeparator Then
                    Call ArrayUtils.Push(result, Mid(str, st, en - st + 1))
                End If
                st = 0
            End If
        End If
    Next

    ' ��؂蕶���������炸�������I������ꍇ
    If st = 0 Then
        Call ArrayUtils.Push(result, Mid(str, en + 1, Len(str) - en))
    Else
        subStr = Mid(str, en + 1, st - en - 1)
        If subStr <> "" Then
            Call ArrayUtils.Push(result, subStr)
        End If
        en = i - 1
        If containsSeparator Then
            Call ArrayUtils.Push(result, Mid(str, st, en - st + 1))
        End If
    End If

    SplitByChars = result

End Function

'------------------------------------------------------------------------------
' FUNCTION : ������𐔎����������؂�Ƃ��ĕ������A���ʂ�z��Ɋi�[���ĕԂ��܂��B
'            ��؂蕶�����A�����Ă������́A�ЂƂ̋�؂蕶����Ƃ��Ĉ����܂��B
'            ��F"abc-123def"�̏ꍇ�́A(abc-, def)���Ԃ����B
'
' PARAMS   : str                 - �Ώە�����
'            [containsSeparator] - ��؂蕶��������ʂɊ܂߂邩 (�ȗ����� False)
'
' RETURN   : �����������ʂ��i�[�����z��
'------------------------------------------------------------------------------
Public Function SplitByDigit(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByDigit = SplitByChars(str, CHARS_DIGIT, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ����������s���������؂�Ƃ��ĕ������A���ʂ�z��Ɋi�[���ĕԂ��܂��B
'            ��؂蕶�����A�����Ă������́A�ЂƂ̋�؂蕶����Ƃ��Ĉ����܂��B
'
' PARAMS   : str                 - �Ώە�����
'            [containsSeparator] - ��؂蕶��������ʂɊ܂߂邩 (�ȗ����� False)
'
' RETURN   : �����������ʂ��i�[�����z��
'------------------------------------------------------------------------------
Public Function SplitByNewline(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByNewline = SplitByChars(str, vbCr & vbLf, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������Ŏn�܂邩�ǂ����𔻒肵�܂��B
'
' PARAMS   : str          - �Ώە�����
'            prefix       - ����������
'            [ignoreCase] - �啶���������𖳎����邩 (�ȗ����� False)
'
' RETURN   : �w�肳�ꂽ������Ŏn�܂�ꍇ�� True
'------------------------------------------------------------------------------
Public Function StartsWith(ByVal str As String, ByVal prefix As String, _
                            Optional ByVal ignoreCase As Boolean = False) As Boolean

    StartsWith = IndexOf(str, prefix, 1, ignoreCase) = 1

End Function

'------------------------------------------------------------------------------
' FUNCTION : �����񂪎w�肳�ꂽ������̂����ꂩ�Ŏn�܂邩�ǂ����𔻒肵�܂��B
'
' PARAMS   : str      - �Ώە�����
'            prefixes - ����������ꗗ
'
' RETURN   : �w�肳�ꂽ������Ŏn�܂�ꍇ�� True
'------------------------------------------------------------------------------
Public Function StartsWithAny(ByVal str As String, ParamArray prefixes() As Variant) As Boolean
    Dim params()    As Variant
    Dim param       As Variant

    params = ArrayUtils.Flatten(0, prefixes)

    StartsWithAny = True

    If Not ArrayUtils.IsEmptyArray(params) Then
        For Each param In params
            If StartsWith(str, param) Then
                Exit Function
            End If
        Next
    End If

    StartsWithAny = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ�����񂩂�擪�Ɩ����̋󔒕�������폜���ĕԂ��܂��B
'            (���p/�S�p�X�y�[�X�A�^�u�A���s�������󔒕����Ƃ݂Ȃ��܂�)
'
' PARAMS   : str - �Ώە�����
'
' RETURN   : �擪�Ɩ����̋󔒕�������폜����������
'------------------------------------------------------------------------------
Public Function Trim(ByVal str As String) As String

    Trim = RTrim(LTrim(str))

End Function
