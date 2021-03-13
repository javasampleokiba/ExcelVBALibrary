Attribute VB_Name = "ArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : 配列操作に関するユーティリティモジュール
'
' NOTE   : 各関数の配列を渡す引数に配列以外を指定した場合、
'          特に明記がない限りエラーを発出します。
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・LangUtils
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された動的配列の指定位置に要素を追加します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 追加対象の配列
'            index - 追加する位置インデックス
'            item  - 追加する要素
'
'------------------------------------------------------------------------------
Public Sub Add(ByRef arr As Variant, ByVal index As Long, ByRef item As Variant)
    Dim i       As Long
    Dim idx     As Long

    If index = 0 And IsEmptyArray(arr) Then
        Call Init(arr, item)
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
'
' FUNCTION : 指定された動的配列に指定された配列を追加します。
'
' PARAMS   : arr      - 追加対象の配列
'            otherArr - 追加する配列
'
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
'
' FUNCTION : 指定した要素が指定した配列に含まれるか判定します。
'
' PARAMS   : arr  - 検索対象の配列
'            item - 検索する要素
'
' RETURN   : 存在する場合は True
'
'------------------------------------------------------------------------------
Public Function Contains(ByRef arr As Variant, ByRef item As Variant) As Boolean

    If -1 < IndexOf(arr, item) Then
        Contains = True
    Else
        Contains = False
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した要素すべてが指定した配列に含まれるか判定します。
'
' PARAMS   : arr   - 検索対象の配列
'            items - 検索する要素の配列
'
' RETURN   : すべての要素が存在する場合は True
'
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
'
' FUNCTION : 指定した要素のいずれかが指定した配列に含まれるか判定します。
'
' PARAMS   : arr   - 検索対象の配列
'            items - 検索する要素の配列
'
' RETURN   : いずれかの要素が存在する場合は True
'
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
'
' FUNCTION : 指定された配列の先頭の要素を取得します。
'
' PARAMS   : arr - 検索対象の配列
'
' RETURN   : 先頭の要素
'
'------------------------------------------------------------------------------
Public Function First(ByRef arr As Variant) As Variant

    If IsObject(arr(LBound(arr))) Then
        Set First = arr(LBound(arr))
    Else
        First = arr(LBound(arr))
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された配列から指定した位置インデックスの要素を取得します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 取得対象の配列
'            index - 取得する位置インデックス
'            [def] - 位置インデックスが不正の場合に返すデフォルト値 (省略時はNull)
'
' RETURN   : 指定した位置インデックスの要素 または デフォルト値
'
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
'
' FUNCTION : 指定した要素が指定した配列内で最初に見つかった位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'
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
'
' FUNCTION : 指定した要素が指定した配列内で見つかったすべての位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'
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
'
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef arr As Variant) As Boolean
    Dim i       As Long

    If Not IsArray(arr) Then
        Call Err.Raise(5)   ' プロシージャの呼び出し、または引数が不正です。
    End If
    
    On Error GoTo ErrHandler
    i = UBound(arr)
    IsEmptyArray = False

    Exit Function

ErrHandler:
    IsEmptyArray = True

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された2つの配列が等しいか判定します。
'
' PARAMS   : arr1 - 比較対象の配列1
'            arr2 - 比較対象の配列2
'
' RETURN   : 指定された2つの配列が等しい場合は true
'
' ERROR    : 引数に配列以外を指定した場合
'
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
'
' FUNCTION : 指定された配列の最後の要素を取得します。
'
' PARAMS   : arr - 検索対象の配列
'
' RETURN   : 最後の要素
'
'------------------------------------------------------------------------------
Public Function Last(ByRef arr As Variant) As Variant

    If IsObject(arr(UBound(arr))) Then
        Set Last = arr(UBound(arr))
    Else
        Last = arr(UBound(arr))
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した要素が指定した配列内で最後に見つかった位置インデックスを
'            返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'
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
'
' FUNCTION : 最後から検索して、指定した要素が指定した配列内で見つかった
'            すべての位置インデックスを返します。
'
' PARAMS   : arr     - 検索対象の配列
'            item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'
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
'
' FUNCTION : 配列の要素数を取得します。
'
' PARAMS   : arr - 取得対象の配列
'
' RETURN   : 配列の要素数
'
'------------------------------------------------------------------------------
Public Function Length(ByRef arr As Variant) As Long

    If IsEmptyArray(arr) Then
        Length = 0
        Exit Function
    End If

    Length = UBound(arr) - LBound(arr) + 1

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された動的配列の最後の要素を削除します。
'
' PARAMS   : arr - 削除対象の配列
'
' RETURN   : 削除した要素
'
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
'
' FUNCTION : 指定された動的配列の最後に要素を追加します。
'
' PARAMS   : arr  - 追加対象の配列
'            item - 追加する要素
'
'------------------------------------------------------------------------------
Public Sub Push(ByRef arr As Variant, ByRef item As Variant)

    If IsEmptyArray(arr) Then
        Call Add(arr, 0, item)
    Else
        Call Add(arr, UBound(arr) + 1, item)
    End If

End Sub

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された配列から指定した要素を削除します。
'
' PARAMS   : arr    - 削除対象の配列
'            item   - 削除する要素
'            [size] - 削除する最大数 (省略すると、すべて削除)
'
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
'
' FUNCTION : 指定された配列から指定したすべての要素を削除します。
'
' PARAMS   : arr    - 削除対象の配列
'            items  - 削除する要素の配列
'
'------------------------------------------------------------------------------
Public Sub RemoveAll(ByRef arr As Variant, ByRef items As Variant)
    Dim item    As Variant

    For Each item In items
        Call Remove(arr, item)
    Next

End Sub

'------------------------------------------------------------------------------
'
' FUNCTION : 配列内の指定した位置の要素を削除します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 削除対象の配列
'            index - 削除する位置インデックス
'
' RETURN   : 削除された要素
'
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
'
' FUNCTION : 指定された配列の指定した位置インデックスに要素を設定します。
'            位置インデックスに負数を指定すると配列の後方からの位置インデックス
'            として検索します。
'
' PARAMS   : arr   - 設定対象の配列
'            index - 設定する位置インデックス
'            item  - 設定する要素
'
' RETURN   : 設定前に指定した位置インデックスにあった要素
'
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
'
' FUNCTION : 指定された動的配列の先頭の要素を削除します。
'
' PARAMS   : arr - 削除対象の配列
'
' RETURN   : 削除した要素
'
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
'
' FUNCTION : 指定された配列の要素を昇順にソートします。
'            すべての要素が不等号演算子による比較が可能であることが前提条件です。
'            CompareToメソッドを実装したオブジェクトの配列も引数にできます。
'
' PARAMS   : arr - 処理対象の配列
'
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
'
' FUNCTION : 指定された動的配列の先頭に要素を追加します。
'
' PARAMS   : arr  - 追加対象の配列
'            item - 追加する要素
'
'------------------------------------------------------------------------------
Public Sub Unshift(ByRef arr As Variant, ByRef item As Variant)

    Call Add(arr, 0, item)

End Sub

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

Private Sub Init(ByRef arr As Variant, ByRef item As Variant)

    ReDim arr(0)
    Call SetAt(arr, 0, item)

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
