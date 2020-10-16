Attribute VB_Name = "ArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : 配列操作に関するユーティリティモジュール
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・LangUtils
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
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
