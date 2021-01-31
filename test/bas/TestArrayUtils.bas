Attribute VB_Name = "TestArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : ArrayUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)

    Debug.Print "=== TestArrayUtils ==="

    Call TestContains
    Call TestContainsAll
    Call TestContainsAny
    Call TestIndexOf
    Call TestIndicesOf
    Call TestIsEmptyArray(arr)
    Call TestIsEqual(arr)
    Call TestLastIndexOf
    Call TestLastIndicesOf
    Call TestLength(arr)
    Call TestSort

End Sub

Private Sub TestContains()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant

    Debug.Print "--- TestContains ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = "A"
    arr(5) = "B"
    arr(6) = "C"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.Contains("ABC", 1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(Not ArrayUtils.Contains(emptyArr, 1), 2)

    ' 要素が見つからない場合
    Call PrintResult(Not ArrayUtils.Contains(arr, 4), 3)

    ' 先頭、中間、末尾の位置で見つかる場合
    Call PrintResult(ArrayUtils.Contains(arr, 1), 4)
    Call PrintResult(ArrayUtils.Contains(arr, "A"), 5)
    Call PrintResult(ArrayUtils.Contains(arr, "C"), 6)

End Sub

Private Sub TestContainsAll()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim items1(0)       As Variant
    Dim items2(2)       As Variant

    Debug.Print "--- TestContainsAll ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = "A"
    arr(5) = "B"
    arr(6) = "C"
    items1(0) = 4
    items2(0) = 4
    items2(1) = "D"
    items2(2) = "E"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.ContainsAll("ABC", items1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    On Error Resume Next
    Call ArrayUtils.ContainsAll(items1, "ABC")
    Call PrintResult(Err.number = 5, 2)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(Not ArrayUtils.ContainsAll(emptyArr, items1), 3)
    Call PrintResult(ArrayUtils.ContainsAll(arr, emptyArr), 4)

    ' 要素が1つも見つからない場合
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items1), 5)
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items2), 6)

    ' 一部の要素が見つかる場合
    items2(0) = 1
    items2(1) = "A"
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items2), 7)

    ' すべての要素が見つかる場合
    items1(0) = 1
    items2(2) = "C"
    Call PrintResult(ArrayUtils.ContainsAll(arr, items1), 8)
    Call PrintResult(ArrayUtils.ContainsAll(arr, items2), 9)

End Sub

Private Sub TestContainsAny()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim items1(0)       As Variant
    Dim items2(2)       As Variant

    Debug.Print "--- TestContainsAny ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = "A"
    arr(5) = "B"
    arr(6) = "C"
    items1(0) = 4
    items2(0) = 4
    items2(1) = "D"
    items2(2) = "E"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.ContainsAny("ABC", items1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    On Error Resume Next
    Call ArrayUtils.ContainsAny(items1, "ABC")
    Call PrintResult(Err.number = 5, 2)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(Not ArrayUtils.ContainsAny(emptyArr, items1), 3)
    Call PrintResult(ArrayUtils.ContainsAny(arr, emptyArr), 4)

    ' 要素が1つも見つからない場合
    Call PrintResult(Not ArrayUtils.ContainsAny(arr, items1), 5)
    Call PrintResult(Not ArrayUtils.ContainsAny(arr, items2), 6)

    ' 一部の要素が見つかる場合
    items2(1) = "A"
    Call PrintResult(ArrayUtils.ContainsAny(arr, items2), 7)

    ' すべての要素が見つかる場合
    items1(0) = 1
    items2(0) = 1
    items2(2) = "C"
    Call PrintResult(ArrayUtils.ContainsAny(arr, items1), 8)
    Call PrintResult(ArrayUtils.ContainsAny(arr, items2), 9)

End Sub

Private Sub TestIndexOf()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant

    Debug.Print "--- TestIndexOf ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = "A"
    arr(5) = "B"
    arr(6) = "C"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.IndexOf("ABC", 1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(ArrayUtils.IndexOf(emptyArr, 1) = -1, 2)

    ' 要素が見つからない場合
    Call PrintResult(ArrayUtils.IndexOf(arr, 4) = -1, 3)
    Call PrintResult(ArrayUtils.IndexOf(arr, 3, 4) = -1, 4)
    Call PrintResult(ArrayUtils.IndexOf(arr, "C", 7) = -1, 5)

    ' 先頭、中間、末尾の位置で見つかる場合
    Call PrintResult(ArrayUtils.IndexOf(arr, 1) = 1, 6)
    Call PrintResult(ArrayUtils.IndexOf(arr, "A") = 4, 7)
    Call PrintResult(ArrayUtils.IndexOf(arr, "C") = 6, 8)
    Call PrintResult(ArrayUtils.IndexOf(arr, 3, 3) = 3, 9)
    Call PrintResult(ArrayUtils.IndexOf(arr, "B", 0) = 5, 10)
    Call PrintResult(ArrayUtils.IndexOf(arr, "C", 6) = 6, 11)

End Sub

Private Sub TestIndicesOf()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim actual          As Variant

    Debug.Print "--- TestIndicesOf ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = "A"
    arr(4) = 1
    arr(5) = "A"
    arr(6) = 1

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.IndicesOf("ABC", 1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    actual = ArrayUtils.IndicesOf(emptyArr, 1)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 2)

    ' 要素が見つからない場合
    actual = ArrayUtils.IndicesOf(arr, 3)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 3)
    actual = ArrayUtils.IndicesOf(arr, 2, 3)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 4)
    actual = ArrayUtils.IndicesOf(arr, 1, 7)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 5)

    ' 先頭、中間、末尾の位置で見つかる場合
    actual = ArrayUtils.IndicesOf(arr, 1)
    Call PrintResult(actual(0) = 1, "6-1")
    Call PrintResult(actual(1) = 4, "6-2")
    Call PrintResult(actual(2) = 6, "6-3")
    Call PrintResult(ArrayUtils.Length(actual) = 3, "6-4")
    
    actual = ArrayUtils.IndicesOf(arr, "A", 0)
    Call PrintResult(actual(0) = 3, "7-1")
    Call PrintResult(actual(1) = 5, "7-2")
    Call PrintResult(ArrayUtils.Length(actual) = 2, "7-3")
    
    actual = ArrayUtils.IndicesOf(arr, 1, 4)
    Call PrintResult(actual(0) = 4, "8-1")
    Call PrintResult(actual(1) = 6, "8-2")
    Call PrintResult(ArrayUtils.Length(actual) = 2, "8-3")
    
    actual = ArrayUtils.IndicesOf(arr, 1, 6)
    Call PrintResult(actual(0) = 6, "9-1")
    Call PrintResult(ArrayUtils.Length(actual) = 1, "9-2")

End Sub

Private Sub TestIsEmptyArray(ByVal arr As Variant)
    Dim testArr()       As Variant

    Debug.Print "--- TestIsEmptyArray ---"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.IsEmptyArray("ABC")
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(ArrayUtils.IsEmptyArray(testArr), 2)
    
    ' 空配列以外の以外
    Call PrintResult(Not ArrayUtils.IsEmptyArray(arr), 3)

End Sub

Private Sub TestIsEqual(ByVal arr As Variant)
    Dim i                   As Integer
    Dim emptyArr1()         As Variant
    Dim emptyArr2()         As Variant
    Dim arr1(0 To 10)       As Variant
    Dim arr2(1 To 10)       As Variant
    Dim arr3(0 To 9)        As Variant
    Dim arr4(0 To 10)       As Variant

    Debug.Print "--- TestIsEqual ---"
    
    arr1(0) = 1
    arr1(1) = 2
    arr1(2) = 3
    arr1(3) = 4
    arr1(4) = 5
    arr1(5) = "A"
    arr1(6) = "B"
    arr1(7) = "C"
    arr1(8) = "D"
    arr1(9) = "E"
    For i = 0 To 10
        arr4(i) = arr1(i)
    Next

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.IsEqual("ABC", arr1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    On Error Resume Next
    Call ArrayUtils.IsEqual(arr1, 123)
    Call PrintResult(Err.number = 5, 2)
    On Error GoTo 0

    Call PrintResult(ArrayUtils.IsEqual(emptyArr1, emptyArr2), 3)
    Call PrintResult(Not ArrayUtils.IsEqual(emptyArr1, arr1), 4)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, emptyArr1), 5)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr2), 6)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr3), 7)
    Call PrintResult(Not ArrayUtils.IsEqual(arr2, arr3), 8)
    Call PrintResult(ArrayUtils.IsEqual(arr1, arr4), 9)
    
    arr4(5) = 6
    
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr4), 10)
    Call PrintResult(ArrayUtils.IsEqual(arr, arr), 11)

End Sub

Private Sub TestLastIndexOf()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant

    Debug.Print "--- TestLastIndexOf ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = "A"
    arr(5) = "B"
    arr(6) = "C"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.LastIndexOf("ABC", 1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(ArrayUtils.LastIndexOf(emptyArr, 1) = -1, 2)

    ' 要素が見つからない場合
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 4) = -1, 3)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, "A", 3) = -1, 4)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 1, 0) = -1, 5)

    ' 先頭、中間、末尾の位置で見つかる場合
    Call PrintResult(ArrayUtils.LastIndexOf(arr, "C") = 6, 6)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 3) = 3, 7)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 1) = 1, 8)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, "A", 4) = 4, 9)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 2, 7) = 2, 10)
    Call PrintResult(ArrayUtils.LastIndexOf(arr, 1, 1) = 1, 11)

End Sub

Private Sub TestLastIndicesOf()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim actual          As Variant

    Debug.Print "--- TestLastIndicesOf ---"

    arr(1) = 1
    arr(2) = 2
    arr(3) = "A"
    arr(4) = 1
    arr(5) = "A"
    arr(6) = 1

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.LastIndicesOf("ABC", 1)
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    actual = ArrayUtils.LastIndicesOf(emptyArr, 1)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 2)

    ' 要素が見つからない場合
    actual = ArrayUtils.LastIndicesOf(arr, 3)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 3)
    actual = ArrayUtils.LastIndicesOf(arr, 2, 1)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 4)
    actual = ArrayUtils.LastIndicesOf(arr, 1, 0)
    Call PrintResult(ArrayUtils.IsEmptyArray(actual), 5)

    ' 先頭、中間、末尾の位置で見つかる場合
    actual = ArrayUtils.LastIndicesOf(arr, 1)
    Call PrintResult(actual(0) = 6, "6-1")
    Call PrintResult(actual(1) = 4, "6-2")
    Call PrintResult(actual(2) = 1, "6-3")
    Call PrintResult(ArrayUtils.Length(actual) = 3, "6-4")
    
    actual = ArrayUtils.LastIndicesOf(arr, "A", 7)
    Call PrintResult(actual(0) = 5, "7-1")
    Call PrintResult(actual(1) = 3, "7-2")
    Call PrintResult(ArrayUtils.Length(actual) = 2, "7-3")
    
    actual = ArrayUtils.LastIndicesOf(arr, 1, 4)
    Call PrintResult(actual(0) = 4, "8-1")
    Call PrintResult(actual(1) = 1, "8-2")
    Call PrintResult(ArrayUtils.Length(actual) = 2, "8-3")
    
    actual = ArrayUtils.LastIndicesOf(arr, 1, 1)
    Call PrintResult(actual(0) = 1, "9-1")
    Call PrintResult(ArrayUtils.Length(actual) = 1, "9-2")

End Sub

Private Sub TestLength(ByVal arr As Variant)
    Dim testArr()       As Variant
    Dim testArr2(0)     As Variant

    Debug.Print "--- TestLength ---"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.Length("ABC")
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    ' 空配列
    Call PrintResult(ArrayUtils.Length(testArr) = 0, 2)

    ' 空配列以外
    Call PrintResult(ArrayUtils.Length(testArr2) = 1, 3)
    Call PrintResult(ArrayUtils.Length(arr) = 18, 4)

End Sub

Private Sub TestSort()
    Dim i               As Long
    Dim arr1(1000)      As Long
    Dim arr2(1000)      As String
    Dim arr3(5)         As MyClass
    Dim result          As Boolean

    Debug.Print "--- TestSort ---"
 
    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.Sort("ABC")
    Call PrintResult(Err.number = 5, 1)
    On Error GoTo 0

    Randomize
 
    ' 1000コのランダムな数値を配列に格納
    For i = 0 To UBound(arr1)
        arr1(i) = Int((Rnd * 1000) + 1)
    Next

    ' 昇順にソート
    Call ArrayUtils.Sort(arr1)

    result = True
    For i = 0 To UBound(arr1) - 1
        If arr1(i) > arr1(i + 1) Then
            result = False
            Exit For
        End If
    Next

    Call PrintResult(result, 2)
 
    ' 1000コのランダムな文字を配列に格納
    For i = 0 To UBound(arr2)
        arr2(i) = Chr(Int((Rnd * 58) + 65))
    Next

    ' 昇順にソート
    Call ArrayUtils.Sort(arr2)

    result = True
    For i = 0 To UBound(arr2) - 1
        If arr2(i) > arr2(i + 1) Then
            result = False
            Exit For
        End If
    Next

    Call PrintResult(result, 3)

    For i = 0 To UBound(arr3)
        Set arr3(i) = New MyClass
    Next
 
    Call arr3(0).SetName("Mark")
    Call arr3(1).SetName("Yoko")
    Call arr3(2).SetName("Jim")
    Call arr3(3).SetName("George")
    Call arr3(4).SetName("David")
    Call arr3(5).SetName("Cindy")

    ' オブジェクト同士を比較しソート
    Call ArrayUtils.Sort(arr3)
 
    result = True
    For i = 0 To UBound(arr3) - 1
        If arr3(i).GetName() > arr3(i + 1).GetName() Then
            result = False
            Exit For
        End If
    Next

    Call PrintResult(result, 4)

End Sub
