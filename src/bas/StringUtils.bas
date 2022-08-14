Attribute VB_Name = "StringUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : 文字列操作に関するユーティリティモジュール
'
' NOTE   : サロゲートペアには対応していません。
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・ArrayUtils
'------------------------------------------------------------------------------

' 空白文字
Private Const CHARS_BLANK As String = " 　" & vbTab & vbCr & vbLf
' 数字
Private Const CHARS_DIGIT As String = "0123456789"
' 英字
Private Const CHARS_ALPHA As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

'------------------------------------------------------------------------------
' FUNCTION : 文字列が指定された文字列で終わらない場合、末尾に付加して返します。
'
' PARAMS   : str          - 対象文字列
'            suffix       - 追加する文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列を付加した文字列
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
' FUNCTION : 文字列に指定された文字列が含まれるか判定します。
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列が含まれる場合は True
'------------------------------------------------------------------------------
Public Function Contains(ByVal str As String, ByVal searchStr As String, _
                        Optional ByVal ignoreCase As Boolean = False) As Boolean

    Contains = 0 < IndexOf(str, searchStr, 1, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列に指定された文字列すべてが含まれるか判定します。
'
' PARAMS   : str        - 対象文字列
'            searchStrs - 検索文字列一覧
'
' RETURN   : 指定された文字列すべてが含まれる場合は True
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
' FUNCTION : 文字列に指定された文字列のいずれかが含まれるか判定します。
'
' PARAMS   : str        - 対象文字列
'            searchStrs - 検索文字列一覧
'
' RETURN   : 指定された文字列のいずれかが含まれる場合は True
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
' FUNCTION : 文字列に指定された文字列がいくつ含まれるかカウントします。
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列が見つかった数
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
' FUNCTION : 文字列が指定された文字列で終わるかどうかを判定します。
'
' PARAMS   : str          - 対象文字列
'            suffix       - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列で終わる場合は True
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
' FUNCTION : 文字列が指定された文字列のいずれかで終わるかどうかを判定します。
'
' PARAMS   : str      - 対象文字列
'            suffixes - 検索文字列一覧
'
' RETURN   : 指定された文字列で終わる場合は True
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
' FUNCTION : 指定された2つの文字列が等しいかを判定します。
'
' PARAMS   : str1         - 判定対象文字列1
'            str2         - 判定対象文字列2
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 2つの文字列が等しい場合は True
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
' FUNCTION : 文字列が指定された文字列のいずれかと等しいかを判定します。
'
' PARAMS   : str  - 判定対象文字列
'            strs - 判定対象文字列一覧
'
' RETURN   : いずれかの文字列が等しい場合は True
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
' FUNCTION : 指定された文字列のうち、空白文字列ではない最初の文字列を返します。
'            (見つからない場合は空文字を返します)
'
' PARAMS   : strs - 候補文字列一覧
'
' RETURN   : 空白文字列ではない文字列
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
' FUNCTION : 文字列内で指定された文字列が最初に現れる位置インデックスを返します。
'            (先頭の場合は1です。見つからない場合は0を返します)
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [start]      - 検索開始位置 (省略時は1)
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列が最初に現れる位置インデックス
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
' FUNCTION : 文字列内で指定された文字列のいずれかが最初に現れる
'            位置インデックスを返します。
'            (先頭の場合は1です。見つからない場合は0を返します)
'
' PARAMS   : str        - 対象文字列
'            searchStrs - 検索文字列一覧
'
' RETURN   : 指定された文字列のいずれかが最初に現れる位置インデックス
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
' FUNCTION : すべての文字列が空白文字のみで構成されているか判定します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : strs - 対象文字列一覧
'
' RETURN   : すべてが空白文字のみで構成されている場合は True
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
' FUNCTION : すべての文字列が空文字か判定します。
'
' PARAMS   : strs - 対象文字列一覧
'
' RETURN   : すべての文字列が空文字の場合は True
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
' FUNCTION : 文字列が英字のみで構成されているか判定します。
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 英字のみで構成されている場合は True
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
' FUNCTION : 文字列が英数字のみで構成されているか判定します。
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 英数字のみで構成されている場合は True
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
' FUNCTION : いずれかの文字列が空白文字のみで構成されているか判定します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : strs - 対象文字列一覧
'
' RETURN   : いずれかが空白文字のみで構成されている場合は True
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
' FUNCTION : いずれかの文字列が空文字か判定します。
'
' PARAMS   : strs - 対象文字列一覧
'
' RETURN   : いずれかの文字列が空文字の場合は True
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
' FUNCTION : 文字列が空白文字のみで構成されているか判定します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 空白文字のみで構成されている場合は True
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
' FUNCTION : 文字列が数字のみで構成されているか判定します。
'            ("0"～"9"を数字とみなします)
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 数字のみで構成されている場合は True
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
' FUNCTION : 文字列内で指定された文字列が最後に現れる位置インデックスを返します。
'            (見つからない場合は0を返します)
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [start]      - 検索開始位置 (省略時は末尾)
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列が最後に現れる位置インデックス
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
' FUNCTION : 文字列内で指定された文字列のいずれかが最後に現れる
'            位置インデックスを返します。
'            (見つからない場合は0を返します)
'
' PARAMS   : str        - 対象文字列
'            searchStrs - 検索文字列一覧
'
' RETURN   : 指定された文字列のいずれかが最後に現れる位置インデックス
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
' FUNCTION : 文字列から指定された文字列より前の部分文字列を返します。
'            (見つからない場合は空文字を返します)
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列より前の部分文字列
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
' FUNCTION : 指定された文字列から先頭の空白文字列を削除して返します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 先頭の空白文字列を削除した文字列
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
' FUNCTION : 文字列から指定された開始・終了文字列の間の部分文字列を返します。
'            (見つからない場合は空文字を返します)
'
' PARAMS   : str          - 対象文字列
'            beforeStr    - 開始文字列
'            afterStr     - 終了文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された開始・終了文字列の間の部分文字列
'------------------------------------------------------------------------------
Public Function MidBetween(ByVal str As String, ByVal beforeStr As String, _
                            ByVal afterStr As String, _
                            Optional ByVal ignoreCase As Boolean = False) As String

    MidBetween = LeftBefore(RightAfter(str, beforeStr, ignoreCase), afterStr, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列の指定位置の文字列を、指定した別の文字列に置き換えます。
'            (指定した範囲が範囲外の場合は何もせずそのまま返します)
'
' PARAMS   : str        - 対象文字列
'            replaceStr - 置換する文字列
'            startIdx   - 検索開始位置
'            endIdx     - 検索終了位置
'
' RETURN   : 置換後の文字列
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
' FUNCTION : 指定された文字列を区切り文字列が最初に現れる位置で3つに分割し、
'            区切りより前、区切り文字列、区切りより後を格納した配列を返します。
'            区切り文字列が見つからない場合は、2～3番目の要素が空文字になります。
'
' PARAMS   : str          - 対象文字列
'            separator    - 区切り文字列
'            [start]      - 検索開始位置 (省略時は1)
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 区切りより前、区切り文字列、区切りより後を格納した配列
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
' FUNCTION : 文字列が指定された文字列で始まらない場合、先頭に付加して返します。
'
' PARAMS   : str          - 対象文字列
'            preffix      - 追加する文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列を付加した文字列
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
' FUNCTION : 文字列から指定した文字列を削除します。
'
' PARAMS   : str          - 対象文字列
'            removeStr    - 削除する文字列
'            [start]      - 検索開始位置 (省略時は1)
'            [size]       - 削除する最大数 (省略、または0を指定するとすべて削除)
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 削除後の文字列
'------------------------------------------------------------------------------
Public Function Remove(ByVal str As String, ByVal removeStr As String, _
                        Optional ByVal start As Long = 1, Optional ByVal size As Long = 0, _
                        Optional ByVal ignoreCase As Boolean = False) As String

    Remove = Replace(str, removeStr, "", start, size, ignoreCase)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列が指定された文字列で終わる場合、その文字列を末尾から取り除きます。
'            (見つからない場合は何もせず返します)
'
' PARAMS   : str          - 対象文字列
'            suffix       - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列を末尾から削除した文字列
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
' FUNCTION : 文字列が指定された文字列で始まる場合、その文字列を先頭から取り除きます。
'            (見つからない場合は何もせず返します)
'
' PARAMS   : str          - 対象文字列
'            prefix       - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列を先頭から削除した文字列
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
' FUNCTION : 文字列の先頭から指定した文字列を検索し、指定した別の文字列に置き換えます。
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            replaceStr   - 置換する文字列
'            [start]      - 検索開始位置 (省略時は1)
'            [size]       - 置換する最大数 (省略、または0を指定するとすべて置換)
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 置換後の文字列
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
' FUNCTION : 指定された文字列の並びを反転します。
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 反転後の文字列
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
' FUNCTION : 文字列から指定された文字列より後の部分文字列を返します。
'            (見つからない場合は空文字を返します)
'
' PARAMS   : str          - 対象文字列
'            searchStr    - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列より後の部分文字列
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
' FUNCTION : 文字列を指定された文字数だけ回転(移動)させます。
'            距離に正の値を指定すると後ろ方向に移動し、
'            負の値を指定すると前方向に移動します。
'
' PARAMS   : str      - 対象文字列
'            distance - 文字数(移動距離)
'
' RETURN   : 回転後の文字列
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
' FUNCTION : 指定された文字列から末尾の空白文字列を削除して返します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 末尾の空白文字列を削除した文字列
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
' FUNCTION : 文字列を英字文字列を区切りとして分割し、結果を配列に格納して返します。
'            区切り文字が連続している限りは、ひとつの区切り文字列として扱います。
'            例："1abc2def3"の場合は、(1, 2, 3)が返される。
'
' PARAMS   : str                 - 対象文字列
'            [containsSeparator] - 区切り文字列も結果に含めるか (省略時は False)
'
' RETURN   : 分割した結果を格納した配列
'------------------------------------------------------------------------------
Public Function SplitByAlpha(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByAlpha = SplitByChars(str, CHARS_ALPHA, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列を空白文字列を区切りとして分割し、結果を配列に格納して返します。
'            区切り文字が連続している限りは、ひとつの区切り文字列として扱います。
'            例："a   b　　c"の場合は、(a, b, c)が返される。
'
' PARAMS   : str                 - 対象文字列
'            [containsSeparator] - 区切り文字列も結果に含めるか (省略時は False)
'
' RETURN   : 分割した結果を格納した配列
'------------------------------------------------------------------------------
Public Function SplitByBlank(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByBlank = SplitByChars(str, CHARS_BLANK, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列を指定したいずれかの文字を区切りとして分割し、
'            結果を配列に格納して返します。
'            区切り文字が連続している限りは、ひとつの区切り文字列として扱います。
'            例："abc-|-|-def"に対して、区切り文字一覧字"-|"を指定した場合、
'                (abc, def)が返される。
'
' PARAMS   : str                 - 対象文字列
'            separateChars       - 区切り文字一覧
'            [containsSeparator] - 区切り文字列も結果に含めるか (省略時は False)
'
' RETURN   : 分割した結果を格納した配列
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
        ' 区切り文字が見つかっていない場合
        If st = 0 Then
            ' 区切り文字の場合
            If 0 < InStr(separateChars, c) Then
                st = i
            End If
        ' 区切り文字が見つかっている場合
        Else
            ' 区切り文字ではない場合
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

    ' 区切り文字が見つからず走査が終わった場合
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
' FUNCTION : 文字列を数字文字列を区切りとして分割し、結果を配列に格納して返します。
'            区切り文字が連続している限りは、ひとつの区切り文字列として扱います。
'            例："abc-123def"の場合は、(abc-, def)が返される。
'
' PARAMS   : str                 - 対象文字列
'            [containsSeparator] - 区切り文字列も結果に含めるか (省略時は False)
'
' RETURN   : 分割した結果を格納した配列
'------------------------------------------------------------------------------
Public Function SplitByDigit(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByDigit = SplitByChars(str, CHARS_DIGIT, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列を改行文字列を区切りとして分割し、結果を配列に格納して返します。
'            区切り文字が連続している限りは、ひとつの区切り文字列として扱います。
'
' PARAMS   : str                 - 対象文字列
'            [containsSeparator] - 区切り文字列も結果に含めるか (省略時は False)
'
' RETURN   : 分割した結果を格納した配列
'------------------------------------------------------------------------------
Public Function SplitByNewline(ByVal str As String, _
                                Optional ByVal containsSeparator As Boolean = False) As String()

    SplitByNewline = SplitByChars(str, vbCr & vbLf, containsSeparator)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列が指定された文字列で始まるかどうかを判定します。
'
' PARAMS   : str          - 対象文字列
'            prefix       - 検索文字列
'            [ignoreCase] - 大文字小文字を無視するか (省略時は False)
'
' RETURN   : 指定された文字列で始まる場合は True
'------------------------------------------------------------------------------
Public Function StartsWith(ByVal str As String, ByVal prefix As String, _
                            Optional ByVal ignoreCase As Boolean = False) As Boolean

    StartsWith = IndexOf(str, prefix, 1, ignoreCase) = 1

End Function

'------------------------------------------------------------------------------
' FUNCTION : 文字列が指定された文字列のいずれかで始まるかどうかを判定します。
'
' PARAMS   : str      - 対象文字列
'            prefixes - 検索文字列一覧
'
' RETURN   : 指定された文字列で始まる場合は True
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
' FUNCTION : 指定された文字列から先頭と末尾の空白文字列を削除して返します。
'            (半角/全角スペース、タブ、改行文字を空白文字とみなします)
'
' PARAMS   : str - 対象文字列
'
' RETURN   : 先頭と末尾の空白文字列を削除した文字列
'------------------------------------------------------------------------------
Public Function Trim(ByVal str As String) As String

    Trim = RTrim(LTrim(str))

End Function
