Attribute VB_Name = "ArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
' MODULE : 配列操作に関するユーティリティモジュール
'
' NOTE   : 各関数の配列を渡す引数に配列以外を指定した場合、
'          特に明記がない限りエラーを発出します。
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・LangUtils
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列の指定位置に要素を追加します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 追加対象の配列
'            index - 追加する位置インデックス
'            item  - 追加する要素
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Sub Add(ByRef arr As Variant, ByVal index As Long, ByRef item As Variant)
    Dim i       As Long
    Dim idx     As Long

    If index = 0 And IsEmptyArray(arr) Then
        ReDim arr(0)
        Call SetAt(arr, 0, item)
        Exit Sub
    End If

    idx = ActualIndex(arr, index)

    Call CheckBoundForInsert(arr, idx)

    Call Resize(arr, 1)
    For i = UBound(arr) - 1 To idx Step -1
        Call SetAt(arr, i + 1, arr(i))
    Next
    Call SetAt(arr, idx, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列に指定された配列を追加します。
'
' PARAMS   : arr      - 追加対象の配列
'            otherArr - 追加する配列
'------------------------------------------------------------------------------
Public Sub Concat(ByRef arr As Variant, ByRef otherArr As Variant)
    Dim idx     As Long
    Dim item    As Variant

    If IsEmptyArray(otherArr) Then Exit Sub

    If IsEmptyArray(arr) Then
        idx = 0
        ReDim Preserve arr(0 To Length(otherArr) - 1)
    Else
        idx = UBound(arr) + 1
        Call Resize(arr, Length(otherArr))
    End If

    For Each item In otherArr
        Call SetAt(arr, idx, item)
        idx = idx + 1
    Next

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素が指定した配列に含まれるか判定します。
'
' PARAMS   : arr  - 検索対象の配列
'            item - 検索する要素
'
' RETURN   : 存在する場合は True
'------------------------------------------------------------------------------
Public Function Contains(ByRef arr As Variant, ByRef item As Variant) As Boolean

    If -1 < IndexOf(arr, item) Then
        Contains = True
    Else
        Contains = False
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素すべてが指定した配列に含まれるか判定します。
'
' PARAMS   : arr   - 検索対象の配列
'            items - 検索する要素の配列
'
' RETURN   : すべての要素が存在する場合は True
'------------------------------------------------------------------------------
Public Function ContainsAll(ByRef arr As Variant, ByRef items As Variant) As Boolean
    Dim item    As Variant

    If IsEmptyArray(items) Then
        ContainsAll = True
        Exit Function
    End If

    ContainsAll = False

    If IsEmptyArray(arr) Then Exit Function

    For Each item In items
        If IndexOf(arr, item) = -1 Then
            Exit Function
        End If
    Next

    ContainsAll = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素のいずれかが指定した配列に含まれるか判定します。
'
' PARAMS   : arr   - 検索対象の配列
'            items - 検索する要素の配列
'
' RETURN   : いずれかの要素が存在する場合は True
'------------------------------------------------------------------------------
Public Function ContainsAny(ByRef arr As Variant, ByRef items As Variant) As Boolean
    Dim item    As Variant

    If IsEmptyArray(items) Then
        ContainsAny = True
        Exit Function
    End If

    If IsEmptyArray(arr) Then
        ContainsAny = False
        Exit Function
    End If

    ContainsAny = True

    For Each item In items
        If -1 < IndexOf(arr, item) Then
            Exit Function
        End If
    Next

    ContainsAny = False

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された要素が指定された配列にいくつ存在するかカウントします。
'
' PARAMS   : arr  - 検索対象の配列
'            item - 検索する要素
'
' RETURN   : 見つかった要素の個数
'------------------------------------------------------------------------------
Public Function Count(ByRef arr As Variant, ByRef item As Variant) As Long
    Dim i       As Long
    Dim tmp     As Variant
    Dim result  As Long

    result = 0

    If IsEmptyArray(arr) Then
        Count = 0
        Exit Function
    End If

    For Each tmp In arr
        If LangUtils.IsEqual(tmp, item) Then
            result = result + 1
        End If
    Next

    Count = result

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した配列の各要素を指定した要素に置き換えます。
'            空の配列が指定された場合は何もしません。
'
' PARAMS   : arr     - 置換対象の配列
'            item    - 置換する要素
'            [start] - 開始位置インデックス
'                      (負数を指定すると後方からの位置インデックスになります)
'                      (省略すると先頭を指定したことになります)
'            [size]  - 置換する個数 (省略すると最後の要素まで置換します)
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Sub Fill(ByRef arr As Variant, ByRef item As Variant, _
                Optional ByVal start As Long = 0, Optional ByVal size As Long = 0)
    Dim i       As Long
    Dim st      As Long
    Dim en      As Long

    If IsEmptyArray(arr) Then
        Exit Sub
    End If

    If start = 0 Then
        start = LBound(arr)
    End If

    st = ActualIndex(arr, start)
    If 0 < size Then
        en = st + size - 1
        If UBound(arr) < en Then
            en = UBound(arr)
        End If
    Else
        en = UBound(arr)
    End If

    Call CheckBound(arr, st)
    Call CheckBound(arr, en)

    For i = st To en
        Call SetAt(arr, i, item)
    Next

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の先頭の要素を取得します。
'
' PARAMS   : arr - 検索対象の配列
'
' RETURN   : 先頭の要素
'------------------------------------------------------------------------------
Public Function First(ByRef arr As Variant) As Variant

    If IsObject(arr(LBound(arr))) Then
        Set First = arr(LBound(arr))
    Else
        First = arr(LBound(arr))
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列から指定した位置インデックスの要素を取得します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 取得対象の配列
'            index - 取得する位置インデックス
'            [def] - 位置インデックスが不正の場合に返すデフォルト値 (省略時はNull)
'
' RETURN   : 指定した位置インデックスの要素 または デフォルト値
'------------------------------------------------------------------------------
Public Function GetAt(ByRef arr As Variant, ByVal index As Long, _
                        Optional ByRef def As Variant = Null) As Variant
    Dim idx     As Long

    If IsObject(def) Then
        Set GetAt = def
    Else
        GetAt = def
    End If

    If IsEmptyArray(arr) Then Exit Function

    idx = ActualIndex(arr, index)

    If LBound(arr) <= idx And idx <= UBound(arr) Then
        If IsObject(arr(idx)) Then
            Set GetAt = arr(idx)
        Else
            GetAt = arr(idx)
        End If
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素が指定した配列内で最初に見つかった位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'------------------------------------------------------------------------------
Public Function IndexOf(ByRef arr As Variant, ByRef item As Variant, _
                        Optional ByVal start As Long = -1) As Long
    Dim i       As Long
    Dim idx     As Long

    If IsEmptyArray(arr) Then
        IndexOf = -1
        Exit Function
    End If

    If start < LBound(arr) Then
        idx = LBound(arr)
    Else
        idx = start
    End If

    For i = idx To UBound(arr)
        If LangUtils.IsEqual(arr(i), item) Then
            IndexOf = i
            Exit Function
        End If
    Next

    IndexOf = -1

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素が指定した配列内で見つかったすべての位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'------------------------------------------------------------------------------
Public Function IndicesOf(ByRef arr As Variant, ByRef item As Variant, _
                          Optional ByVal start As Long = -1) As Long()
    Dim i               As Long
    Dim idx             As Long
    Dim cnt             As Long
    Dim result()        As Long
    Dim emptyArr()      As Long

    If IsEmptyArray(arr) Then
        IndicesOf = emptyArr
        Exit Function
    End If

    If start < LBound(arr) Then
        idx = LBound(arr)
    Else
        idx = start
    End If

    cnt = 0
    ReDim result(Length(arr) - 1)
    For i = idx To UBound(arr)
        If LangUtils.IsEqual(arr(i), item) Then
            result(cnt) = i
            cnt = cnt + 1
        End If
    Next

    If 0 < cnt Then
        ReDim Preserve result(cnt - 1)
        IndicesOf = result
    Else
        IndicesOf = emptyArr
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列が空であるか判定します。
'            初期化されていない、あるいはErase実行後の動的配列の場合に
'            true を返します。
'
' PARAMS   : arr - 判定対象の配列
'
' RETURN   : 指定された配列が空である場合は true
'
' ERROR    : 引数に配列以外を指定した場合
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef arr As Variant) As Boolean

    If Not IsArray(arr) Then
        Call Err.Raise(5)   ' プロシージャの呼び出し、または引数が不正です。
    End If

    On Error GoTo ErrHandler
    If UBound(arr) < 0 Then
        IsEmptyArray = True
    Else
        IsEmptyArray = False
    End If

    Exit Function

ErrHandler:
    IsEmptyArray = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された2つの配列が等しいか判定します。
'
' PARAMS   : arr1 - 比較対象の配列1
'            arr2 - 比較対象の配列2
'
' RETURN   : 指定された2つの配列が等しい場合は true
'
' ERROR    : 引数に配列以外を指定した場合
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    Dim i       As Long

    IsEqual = False

    ' 両方とも要素数が設定されていない動的配列の場合
    If IsEmptyArray(arr1) And IsEmptyArray(arr2) Then
        IsEqual = True
        Exit Function
        
    ' いずれかが要素数が設定されていない動的配列の場合
    ElseIf IsEmptyArray(arr1) Or IsEmptyArray(arr2) Then
        Exit Function
    End If

    ' 配列の最小・最大インデックスが等しくない場合
    If LBound(arr1) <> LBound(arr2) Then Exit Function
    If UBound(arr1) <> UBound(arr2) Then Exit Function

    ' 全要素を比較
    For i = LBound(arr1) To UBound(arr1)
        If Not LangUtils.IsEqual(arr1(i), arr2(i)) Then
            Exit Function
        End If
    Next

    IsEqual = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の各要素の文字列表現を指定された区切り文字で連結し、
'            返します。
'
' PARAMS   : arr   - 連結対象の配列
'            [sep] - 区切り文字 (省略された場合は空文字)
'
' RETURN   : 連結された文字列
'------------------------------------------------------------------------------
Public Function Join(ByRef arr As Variant, Optional ByVal sep As String = "") As String
    Dim i       As Long
    Dim result  As String

    result = ""

    If IsEmptyArray(arr) Then
        Join = result
        Exit Function
    End If

    result = LangUtils.ToString(arr(LBound(arr)))
    For i = LBound(arr) + 1 To UBound(arr)
        result = result & sep & LangUtils.ToString(arr(i))
    Next

    Join = result

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の最後の要素を取得します。
'
' PARAMS   : arr - 検索対象の配列
'
' RETURN   : 最後の要素
'------------------------------------------------------------------------------
Public Function Last(ByRef arr As Variant) As Variant

    If IsObject(arr(UBound(arr))) Then
        Set Last = arr(UBound(arr))
    Else
        Last = arr(UBound(arr))
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素が指定した配列内で最後に見つかった位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'------------------------------------------------------------------------------
Public Function LastIndexOf(ByRef arr As Variant, ByRef item As Variant, _
                            Optional ByVal start As Long = -1) As Long
    Dim i       As Long
    Dim idx     As Long

    If IsEmptyArray(arr) Then
        LastIndexOf = -1
        Exit Function
    End If

    If start < 0 Or UBound(arr) < start Then
        idx = UBound(arr)
    Else
        idx = start
    End If

    For i = idx To LBound(arr) Step -1
        If LangUtils.IsEqual(arr(i), item) Then
            LastIndexOf = i
            Exit Function
        End If
    Next

    LastIndexOf = -1

End Function

'------------------------------------------------------------------------------
' FUNCTION : 最後から検索して、指定した要素が指定した配列内で見つかった
'            すべての位置インデックスを返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'------------------------------------------------------------------------------
Public Function LastIndicesOf(ByRef arr As Variant, ByRef item As Variant, _
                              Optional ByVal start As Long = -1) As Long()
    Dim i               As Long
    Dim idx             As Long
    Dim cnt             As Long
    Dim result()        As Long
    Dim emptyArr()      As Long

    If IsEmptyArray(arr) Then
        LastIndicesOf = result
        Exit Function
    End If

    If start < 0 Or UBound(arr) < start Then
        idx = UBound(arr)
    Else
        idx = start
    End If

    cnt = 0
    ReDim result(Length(arr) - 1)
    For i = idx To LBound(arr) Step -1
        If LangUtils.IsEqual(arr(i), item) Then
            result(cnt) = i
            cnt = cnt + 1
        End If
    Next

    If 0 < cnt Then
        ReDim Preserve result(cnt - 1)
        LastIndicesOf = result
    Else
        LastIndicesOf = emptyArr
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 配列の要素数を取得します。
'
' PARAMS   : arr - 取得対象の配列
'
' RETURN   : 配列の要素数
'------------------------------------------------------------------------------
Public Function Length(ByRef arr As Variant) As Long

    If IsEmptyArray(arr) Then
        Length = 0
        Exit Function
    End If

    Length = UBound(arr) - LBound(arr) + 1

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列の最後の要素を削除します。
'
' PARAMS   : arr - 削除対象の配列
'
' RETURN   : 削除した要素
'------------------------------------------------------------------------------
Public Function Pop(ByRef arr As Variant) As Variant
    Dim i       As Long
    Dim uIdx    As Long

    uIdx = UBound(arr)
    If IsObject(arr(uIdx)) Then
        Set Pop = arr(uIdx)
    Else
        Pop = arr(uIdx)
    End If

    If Length(arr) = 1 Then
        Erase arr
    Else
        Call Resize(arr, -1)
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列の最後に要素を追加します。
'
' PARAMS   : arr  - 追加対象の配列
'            item - 追加する要素
'------------------------------------------------------------------------------
Public Sub Push(ByRef arr As Variant, ByRef item As Variant)

    If IsEmptyArray(arr) Then
        Call Add(arr, 0, item)
    Else
        Call Add(arr, UBound(arr) + 1, item)
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列から指定した要素を削除します。
'
' PARAMS   : arr    - 削除対象の配列
'            item   - 削除する要素
'            [size] - 削除する最大数 (省略すると、すべて削除)
'------------------------------------------------------------------------------
Public Sub Remove(ByRef arr As Variant, ByRef item As Variant, _
                  Optional ByVal size As Long = 0)
    Dim idx     As Long
    Dim cur     As Long
    Dim cnt     As Long

    cur = -1
    cnt = 0

    Do While True
        idx = IndexOf(arr, item, cur)
        If 0 <= idx Then
            Call RemoveAt(arr, idx)
            cnt = cnt + 1

            ' 削除最大数に達したら終了
            If 0 < size And cnt = size Then
                Exit Sub
            End If

        ' 見つからなければ終了
        Else
            Exit Sub
        End If

        cur = idx
    Loop

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列から指定したすべての要素を削除します。
'
' PARAMS   : arr    - 削除対象の配列
'            items  - 削除する要素の配列
'------------------------------------------------------------------------------
Public Sub RemoveAll(ByRef arr As Variant, ByRef items As Variant)
    Dim item    As Variant

    For Each item In items
        Call Remove(arr, item)
    Next

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 配列内の指定した位置の要素を削除します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 削除対象の配列
'            index - 削除する位置インデックス
'
' RETURN   : 削除された要素
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Function RemoveAt(ByRef arr As Variant, ByVal index As Long) As Variant
    Dim i       As Long
    Dim idx     As Long

    idx = ActualIndex(arr, index)

    Call CheckBound(arr, idx)

    If IsObject(arr(idx)) Then
        Set RemoveAt = arr(idx)
    Else
        RemoveAt = arr(idx)
    End If

    For i = idx To UBound(arr) - 1
        Call SetAt(arr, i, arr(i + 1))
    Next
    Call Resize(arr, -1)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 配列の先頭から指定した要素を検索し、指定した別の要素に置き換えます。
'
' PARAMS   : arr         - 置換対象の配列
'            item        - 置換対象要素
'            replacement - 置換する要素
'            [size]      - 置換する最大数 (省略すると、すべて置換)
'
' RETURN   : 置換した要素数
'------------------------------------------------------------------------------
Public Function Replace(ByRef arr As Variant, ByRef item As Variant, _
                        ByRef replacement As Variant, _
                        Optional ByVal size As Long = 0) As Long
    Dim i       As Long
    Dim cnt     As Long

    If IsEmptyArray(arr) Then
        Replace = 0
        Exit Function
    End If

    cnt = 0
    For i = LBound(arr) To UBound(arr)
        If LangUtils.IsEqual(arr(i), item) Then
            Call SetAt(arr, i, replacement)
            cnt = cnt + 1

            If cnt = size Then
                Exit For
            End If
        End If
    Next

    Replace = cnt

End Function

'------------------------------------------------------------------------------
' FUNCTION : 配列の末尾から指定した要素を検索し、指定した別の要素に置き換えます。
'
' PARAMS   : arr         - 置換対象の配列
'            item        - 置換対象要素
'            replacement - 置換する要素
'            [size]      - 置換する最大数 (省略すると、すべて置換)
'
' RETURN   : 置換した要素数
'------------------------------------------------------------------------------
Public Function ReplaceLast(ByRef arr As Variant, ByRef item As Variant, _
                            ByRef replacement As Variant, _
                            Optional ByVal size As Long = 0) As Long
    Dim i       As Long
    Dim cnt     As Long

    If IsEmptyArray(arr) Then
        ReplaceLast = 0
        Exit Function
    End If

    cnt = 0
    For i = UBound(arr) To LBound(arr) Step -1
        If LangUtils.IsEqual(arr(i), item) Then
            Call SetAt(arr, i, replacement)
            cnt = cnt + 1

            If cnt = size Then
                Exit For
            End If
        End If
    Next

    ReplaceLast = cnt

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の要素の順番を反転します。
'
' PARAMS   : arr - 処理対象の配列
'------------------------------------------------------------------------------
Public Sub Reverse(ByRef arr As Variant)
    Dim i       As Long
    Dim l       As Long
    Dim m       As Long
    Dim lIdx    As Long
    Dim uIdx    As Long

    l = Length(arr)

    If l <= 1 Then
        Exit Sub
    End If

    m = l / 2 - 1
    lIdx = LBound(arr)
    uIdx = UBound(arr)
    For i = lIdx To lIdx + m
        Call Swap(arr, i, uIdx - i + lIdx)
    Next

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 配列の要素を指定された距離だけ回転させます。
'            距離に正の値を指定すると要素は配列の後ろ方向に移動し、
'            負の値を指定すると前方向に移動します。
'
' PARAMS   : arr      - 処理対象の配列
'            distance - 要素の移動距離
'------------------------------------------------------------------------------
Public Sub Rotate(ByRef arr As Variant, ByVal distance As Long)
    Dim i       As Long
    Dim l       As Long
    Dim m       As Long
    Dim idx     As Long
    Dim lIdx    As Long
    Dim uIdx    As Long

    If IsEmptyArray(arr) Then
        Exit Sub
    End If

    l = Length(arr)

    idx = -distance Mod l
    If idx = 0 Then
        Exit Sub
    End If

    If idx < 0 Then
        idx = idx + l
    End If

    m = idx / 2 - 1
    lIdx = LBound(arr)
    uIdx = lIdx + idx - 1
    For i = lIdx To lIdx + m
        Call Swap(arr, i, uIdx - i + lIdx)
    Next

    m = (l - idx) / 2 - 1
    lIdx = LBound(arr) + idx
    uIdx = UBound(arr)
    For i = lIdx To lIdx + m
        Call Swap(arr, i, uIdx - i + lIdx)
    Next

    Call Reverse(arr)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の要素を指定個数ランダムに返します。
'
' PARAMS   : arr    - 取得対象の配列
'            [size] - 取得する個数 (省略すると1つ。最大は配列サイズ)
'            [uniq] - 同じ位置の要素を複数取得不可とするかのフラグ
'
' RETURN   : 選ばれた要素から成る新しい配列
'------------------------------------------------------------------------------
Public Function Sample(ByRef arr As Variant, Optional ByVal size As Long = 1, _
                       Optional ByVal uniq As Boolean = True) As Variant
    Dim i           As Long
    Dim idx         As Long
    Dim lIdx        As Long
    Dim uIdx        As Long
    Dim result()    As Variant
    Dim indices()   As Long

    Randomize

    If Length(arr) < size Then
        size = Length(arr)
    End If

    ReDim result(size - 1)

    If uniq Then
        ReDim indices(Length(arr) - 1)
        uIdx = UBound(indices)
        For i = 0 To uIdx
            indices(i) = i
        Next

        lIdx = LBound(arr)
        For i = 0 To size - 1
            idx = Int(Rnd * (uIdx + 1))
            Call SetAt(result, i, arr(lIdx + indices(idx)))
            indices(idx) = indices(uIdx)
            uIdx = uIdx - 1
        Next
    Else
        lIdx = LBound(arr)
        uIdx = UBound(arr)
        For i = 0 To size - 1
            idx = Int(Rnd * uIdx) + lIdx
            Call SetAt(result, i, arr(idx))
        Next
    End If

    Sample = result

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の指定した位置インデックスに要素を設定します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 設定対象の配列
'            index - 設定する位置インデックス
'            item  - 設定する要素
'
' RETURN   : 設定前に指定した位置インデックスにあった要素
'------------------------------------------------------------------------------
Public Function SetAt(ByRef arr As Variant, ByVal index As Long, ByRef item As Variant) As Variant
    Dim idx     As Long

    idx = ActualIndex(arr, index)

    If IsObject(arr(idx)) Then
        Set SetAt = arr(idx)
    Else
        SetAt = arr(idx)
    End If

    If IsObject(item) Then
        Set arr(idx) = item
    Else
        arr(idx) = item
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列の先頭の要素を削除します。
'
' PARAMS   : arr - 削除対象の配列
'
' RETURN   : 削除した要素
'------------------------------------------------------------------------------
Public Function Shift(ByRef arr As Variant) As Variant
    Dim lIdx    As Long

    lIdx = LBound(arr)
    If IsObject(arr(lIdx)) Then
        Set Shift = arr(lIdx)
    Else
        Shift = arr(lIdx)
    End If

    Call RemoveAt(arr, lIdx)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の要素の順番をランダムに入れ替えます。
'
' PARAMS   : arr - 処理対象の配列
'------------------------------------------------------------------------------
Public Sub Shuffle(ByRef arr As Variant)
    Dim i       As Long
    Dim j       As Long
    Dim idx     As Long
    Dim cnt     As Long

    If Length(arr) <= 1 Then
        Exit Sub
    End If

    Randomize

    cnt = Length(arr)
    For i = LBound(arr) To UBound(arr)
        idx = Int(Rnd * cnt)
        If idx <> 0 Then
            Call Swap(arr, i, i + idx)
        End If
        cnt = cnt - 1
    Next

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定した配列から引数で指定した範囲の要素を取り出し、
'            配列にして返します。
'
' PARAMS   : arr        - 取得対象の配列
'            [startIdx] - 取得する開始位置インデックス (省略すると先頭)
'            [stopIdx]  - 取得する終了位置インデックス。ただし、stopIdxの位置は
'                         含まれない。 (省略すると末尾まで取得)
'
' RETURN   : 取り出した要素から成る新しい配列
'------------------------------------------------------------------------------
Public Function Slice(ByRef arr As Variant, Optional ByVal startIdx As Long = 0, _
                       Optional ByVal stopIdx As Long = 0) As Variant
    Dim i           As Long
    Dim st          As Long
    Dim sp          As Long
    Dim indices()   As Long

    If startIdx < LBound(arr) Then
        st = LBound(arr)
    Else
        st = startIdx
    End If

    If stopIdx <= 0 Or UBound(arr) < stopIdx Then
        sp = UBound(arr) + 1
    Else
        sp = stopIdx
    End If

    If sp <= st Then
        Slice = ValuesAt(arr)
        Exit Function
    End If

    ReDim indices(st To sp - 1)
    For i = st To sp - 1
        indices(i) = i
    Next

    Slice = ValuesAt(arr, indices)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された配列の要素を昇順にソートします。
'            すべての要素が不等号演算子による比較が可能であることが前提条件です。
'            CompareToメソッドを実装したオブジェクトの配列も引数にできます。
'
' PARAMS   : arr - 処理対象の配列
'------------------------------------------------------------------------------
Public Sub Sort(ByRef arr As Variant)

    If Length(arr) <= 1 Then
        Exit Sub
    End If

    If IsObject(arr(LBound(arr))) Then
        Call QuickObjSort(arr, LBound(arr), UBound(arr))
    Else
        Call QuickSort(arr, LBound(arr), UBound(arr))
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列から要素を取り除きつつ、新しい要素を追加します。
'            追加する要素を指定しない場合は単に要素を削除します。
'            削除数を0にすると単に要素を挿入します。
'
' PARAMS   : arr     - 追加対象の配列
'            index   - 配列を変化させ始めるインデックス
'                      (負数を指定すると配列の後方からの位置インデックス)
'            [size]  - 削除する最大要素数 (省略すると開始位置から後ろの全要素を削除)
'            [items] - 追加する要素の配列
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Sub Splice(ByRef arr As Variant, ByVal index As Long, _
                  Optional ByVal size As Long = -1, _
                  Optional ByRef items As Variant = Null)
    Dim i       As Long
    Dim idx     As Long
    Dim cnt     As Long
    Dim diff    As Long
    Dim delSize As Long
    Dim addSize As Long

    ' 追加対象の配列が空、かつ先頭に追加する場合
    If index = 0 And IsEmptyArray(arr) Then
        ' 追加する配列が空ではない場合、連結
        If IsArray(items) Then
            If Not IsEmptyArray(items) Then
                Call Concat(arr, items)
            End If
        End If

        Exit Sub
    End If

    idx = ActualIndex(arr, index)

    Call CheckBoundForInsert(arr, idx)

    ' 削除する要素数計算
    delSize = UBound(arr) - idx + 1
    If 0 <= size And size < delSize Then
        delSize = size
    End If

    ' 追加する配列が指定されている場合
    If IsArray(items) Then
        ' 追加する配列が空ではない場合
        If Not IsEmptyArray(items) Then
            addSize = Length(items)
            cnt = 0

            ' 削除数より追加数の方が多い場合
            If delSize < addSize Then
                diff = addSize - delSize

                Call Resize(arr, diff)

                For i = idx To idx + delSize - 1
                    Call SetAt(arr, i, items(LBound(items) + cnt))
                    cnt = cnt + 1
                Next
                For i = UBound(arr) - diff To idx + delSize Step -1
                    Call SetAt(arr, i + diff, arr(i))
                Next
                For i = idx + delSize To idx + addSize - 1
                    Call SetAt(arr, i, items(LBound(items) + cnt))
                    cnt = cnt + 1
                Next

            ' 削除数が追加数以上の場合
            Else
                diff = delSize - addSize

                For i = idx To idx + addSize - 1
                    Call SetAt(arr, i, items(LBound(items) + cnt))
                    cnt = cnt + 1
                Next
                For i = idx + addSize To UBound(arr) - diff
                    Call SetAt(arr, i, arr(i + diff))
                Next

                Call Resize(arr, -diff)
            End If

            Exit Sub
        End If
    End If

    If 0 < delSize Then
        For i = idx + delSize To UBound(arr)
            Call SetAt(arr, i - delSize, arr(i))
        Next

        Call Resize(arr, -delSize)
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 2つの要素の位置を入れ替えます。
'
' PARAMS   : arr    - 変更対象の配列
'            index1 - 変更する位置インデックス1
'            index2 - 変更する位置インデックス2
'------------------------------------------------------------------------------
Public Sub Swap(ByRef arr As Variant, ByVal index1 As Long, ByVal index2 As Long)
    Dim tmp     As Variant

    If IsObject(arr(index1)) Then
        Set tmp = arr(index1)
    Else
        tmp = arr(index1)
    End If

    Call SetAt(arr, index1, arr(index2))
    Call SetAt(arr, index2, tmp)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列から重複する要素をすべて削除します。
'
' PARAMS   : arr - 削除対象の配列
'------------------------------------------------------------------------------
Public Sub Unique(ByRef arr As Variant)
    Dim i       As Long
    Dim j       As Long
    Dim lIdx    As Long
    Dim found   As Boolean
    Dim cnt     As Long
    Dim tmp()   As Variant

    If Length(arr) <= 1 Then
        Exit Sub
    End If

    lIdx = LBound(arr)
    tmp = arr
    cnt = 1

    For i = lIdx + 1 To UBound(tmp)
        found = False

        For j = lIdx To i - 1
            If LangUtils.IsEqual(tmp(j), tmp(i)) Then
                found = True
                Exit For
            End If
        Next

        If Not found Then
            Call SetAt(arr, lIdx + cnt, tmp(i))
            cnt = cnt + 1
        End If
    Next

    ReDim Preserve arr(lIdx To lIdx + cnt - 1)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定された動的配列の先頭に要素を追加します。
'
' PARAMS   : arr  - 追加対象の配列
'            item - 追加する要素
'------------------------------------------------------------------------------
Public Sub Unshift(ByRef arr As Variant, ByRef item As Variant)

    Call Add(arr, 0, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 指定した配列から引数で指定した位置の要素を取り出し、
'            配列にして返します。
'
' PARAMS   : arr     - 取得対象配列
'            indices - 取得する位置インデックス、または位置インデックスの配列
'
' RETURN   : 取り出した要素から成る配列
'------------------------------------------------------------------------------
Public Function ValuesAt(ByRef arr As Variant, ParamArray indices()) As Variant
    Dim i           As Long
    Dim j           As Long
    Dim l           As Long
    Dim cnt         As Long
    Dim result()    As Variant

    ' 引数の長さを確認
    l = 0
    For i = 0 To UBound(indices)
        If IsArray(indices(i)) Then
            l = l + Length(indices(i))
        Else
            l = l + 1
        End If
    Next

    ' 引数がない場合
    If l = 0 Then
        ValuesAt = result
        Exit Function
    End If

    ReDim result(l - 1)

    cnt = 0
    For i = 0 To UBound(indices)
        If IsArray(indices(i)) Then
            If Not IsEmptyArray(indices(i)) Then
                For j = LBound(indices(i)) To UBound(indices(i))
                    Call SetAt(result, cnt, arr(indices(i)(j)))
                    cnt = cnt + 1
                Next
            End If
        Else
            Call SetAt(result, cnt, arr(indices(i)))
            cnt = cnt + 1
        End If
    Next

    ValuesAt = result

End Function

Private Function ActualIndex(ByRef arr As Variant, ByVal index As Long) As Long

    If 0 <= index Then
        ActualIndex = index
    Else
        ActualIndex = UBound(arr) + index + 1
    End If

End Function

Private Sub CheckBound(ByRef arr As Variant, ByVal index As Long)

    If index < LBound(arr) Or UBound(arr) < index Then
        Call Err.Raise(9)
    End If

End Sub

Private Sub CheckBoundForInsert(ByRef arr As Variant, ByVal index As Long)

    If index < LBound(arr) Or UBound(arr) + 1 < index Then
        Call Err.Raise(9)
    End If

End Sub

Private Sub QuickSort(ByRef arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i       As Long
    Dim j       As Long
    Dim pivot   As Variant
    Dim tmp     As Variant

    If left >= right Then
        Exit Sub
    End If

    i = left
    j = right
    pivot = arr((i + j) / 2)

    Do
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While pivot < arr(j)
            j = j - 1
        Loop
        If j <= i Then Exit Do
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        i = i + 1
        j = j - 1
    Loop

    Call QuickSort(arr, left, i - 1)
    Call QuickSort(arr, j + 1, right)

End Sub

Private Sub QuickObjSort(ByRef arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i       As Long
    Dim j       As Long
    Dim pivot   As Variant
    Dim tmp     As Variant

    If left >= right Then
        Exit Sub
    End If

    i = left
    j = right
    Set pivot = arr((i + j) / 2)

    Do
        Do While 0 < pivot.CompareTo(arr(i))
            i = i + 1
        Loop
        Do While pivot.CompareTo(arr(j)) < 0
            j = j - 1
        Loop
        If j <= i Then Exit Do
        Set tmp = arr(i)
        Set arr(i) = arr(j)
        Set arr(j) = tmp
        i = i + 1
        j = j - 1
    Loop

    Call QuickObjSort(arr, left, i - 1)
    Call QuickObjSort(arr, j + 1, right)

End Sub

Private Sub Resize(ByRef arr As Variant, ByVal size As Long)

    If LBound(arr) <= UBound(arr) + size Then
        ReDim Preserve arr(LBound(arr) To UBound(arr) + size)
    Else
        Erase arr
    End If

End Sub
