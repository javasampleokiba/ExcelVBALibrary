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
    Dim i       As Long

    If IsEmptyArray(items) Then
        ContainsAll = True
        Exit Function
    End If

    ContainsAll = False

    If IsEmptyArray(arr) Then Exit Function

    For i = LBound(items) To UBound(items)
        If IndexOf(arr, items(i)) = -1 Then
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
    Dim i       As Long

    If IsEmptyArray(items) Then
        ContainsAny = True
        Exit Function
    End If

    If IsEmptyArray(arr) Then
        ContainsAny = False
        Exit Function
    End If

    ContainsAny = True

    For i = LBound(items) To UBound(items)
        If -1 < IndexOf(arr, items(i)) Then
            Exit Function
        End If
    Next

    ContainsAny = False

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
