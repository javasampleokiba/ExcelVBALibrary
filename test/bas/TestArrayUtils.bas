Attribute VB_Name = "TestArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : ArrayUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)

    Debug.Print "=== TestArrayUtils ==="

    Call TestAdd
    Call TestCBoolArray
    Call TestCByteArray
    Call TestCCurArray
    Call TestCDateArray
    Call TestCDblArray
    Call TestCIntArray
    Call TestCLngArray
    Call TestCSngArray
    Call TestCStrArray
    Call TestConcat
    Call TestContains
    Call TestContainsAll
    Call TestContainsAny
    Call TestFill
    Call TestFirst
    Call TestFlatten
    Call TestGetAt
    Call TestIndexOf
    Call TestIndicesOf
    Call TestIsEmptyArray(arr)
    Call TestIsEqual(arr)
    Call TestJoin(arr)
    Call TestLast
    Call TestLastIndexOf
    Call TestLastIndicesOf
    Call TestLength(arr)
    Call TestMax
    Call TestMin
    Call TestPop
    Call TestPush
    Call TestRemove
    Call TestRemoveAll
    Call TestRemoveAt
    Call TestReplace
    Call TestReplaceLast
    Call TestReverse
    Call TestRotate
    Call TestSample
    Call TestSetAt
    Call TestShift
    Call TestShuffle
    Call TestSlice
    Call TestSort
    Call TestSplice
    Call TestSwap
    Call TestUnique
    Call TestUnshift
    Call TestValuesAt

End Sub

Private Sub TestAdd()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestInsert ---"

    ' 配列インデックス範囲外の場合
    On Error Resume Next
    Call ArrayUtils.Add(arr, 4, "A")
    Call PrintResult(Err.Number = 9, 1)
    On Error GoTo 0

    ' 空配列に追加
    Call ArrayUtils.Add(emptyArr, 0, "A")
    Call PrintResult(LangUtils.ToString(emptyArr) = "(A)", 2)

    arr = Array(1, 2, 3)

    ' 先頭に追加：(1, 2, 3) => (A, 1, 2, 3)
    Call ArrayUtils.Add(arr, 0, "A")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 1, 2, 3)", 3)

    ' 途中に追加：(A, 1, 2, 3) => (A, 1, B, 2, 3)
    Call ArrayUtils.Add(arr, 2, "B")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 1, B, 2, 3)", 4)

    ' 末尾に追加：(A, 1, B, 2, 3) => (A, 1, B, 2, 3, C)
    Call ArrayUtils.Add(arr, 5, "C")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 1, B, 2, 3, C)", 5)

    arr = Array(1, 2, 3)

    ' 位置インデックスに負数指定(-1)：(1, 2, 3) => (1, 2, A, 3)
    Call ArrayUtils.Add(arr, -1, "A")
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, A, 3)", 6)

    ' 負数で先頭に追加：(1, 2, A, 3) => (B, 1, 2, A, 3)
    Call ArrayUtils.Add(arr, -4, "B")
    Call PrintResult(LangUtils.ToString(arr) = "(B, 1, 2, A, 3)", 7)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Call ArrayUtils.Add(arr, 1, obj)
    Call PrintResult(arr(1) Is obj, 8)

End Sub

Private Sub TestCBoolArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Boolean

    Debug.Print "--- TestCBoolArray ---"

    actual = ArrayUtils.CBoolArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CBoolArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CBoolArray(True, False, "true", 0)
    Call PrintResult(LangUtils.ToString(actual) = "(True, False, True, False)", 3)

    actual = ArrayUtils.CBoolArray(True, Array(False, Array(True, Array(False, emptyArr, False))))
    Call PrintResult(LangUtils.ToString(actual) = "(True, False, True, False, False)", 4)

End Sub

Private Sub TestCByteArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Byte

    Debug.Print "--- TestCByteArray ---"

    actual = ArrayUtils.CByteArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CByteArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CByteArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4)", 3)

    actual = ArrayUtils.CByteArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCCurArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Currency

    Debug.Print "--- TestCCurArray ---"

    actual = ArrayUtils.CCurArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CCurArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CCurArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4.1)", 3)

    actual = ArrayUtils.CCurArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCDateArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Date

    Debug.Print "--- TestCDateArray ---"

    actual = ArrayUtils.CDateArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CDateArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CDateArray("2021/4/10", 40000.001)
    Call PrintResult(LangUtils.ToString(actual) = "(2021/04/10, 2009/07/06 0:01:26)", 3)

    actual = ArrayUtils.CDateArray("2021/4/1", Array("2021/4/2", Array("2021/4/3", Array("2021/4/4", emptyArr, "2021/4/5"))))
    Call PrintResult(LangUtils.ToString(actual) = "(2021/04/01, 2021/04/02, 2021/04/03, 2021/04/04, 2021/04/05)", 4)

End Sub

Private Sub TestCDblArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Double

    Debug.Print "--- TestCDblArray ---"

    actual = ArrayUtils.CDblArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CDblArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CDblArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4.1)", 3)

    actual = ArrayUtils.CDblArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCIntArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Integer

    Debug.Print "--- TestCIntArray ---"

    actual = ArrayUtils.CIntArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CIntArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CIntArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4)", 3)

    actual = ArrayUtils.CIntArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCLngArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Long

    Debug.Print "--- TestCLngArray ---"

    actual = ArrayUtils.CLngArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CLngArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CLngArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4)", 3)

    actual = ArrayUtils.CLngArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCSngArray()
    Dim emptyArr()      As Variant
    Dim actual()        As Single

    Debug.Print "--- TestCSngArray ---"

    actual = ArrayUtils.CSngArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    On Error Resume Next
    actual = ArrayUtils.CSngArray("A")
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    actual = ArrayUtils.CSngArray(1, 2, "3", "4.1")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4.1)", 3)

    actual = ArrayUtils.CSngArray(1, Array(2, Array(3, Array(4, emptyArr, 5))))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, 4, 5)", 4)

End Sub

Private Sub TestCStrArray()
    Dim emptyArr()      As Variant
    Dim actual()        As String

    Debug.Print "--- TestCStrArray ---"

    actual = ArrayUtils.CStrArray(emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    actual = ArrayUtils.CStrArray(1, "A", True, ActiveSheet, Null, Empty)
    Call PrintResult(LangUtils.ToString(actual) = "(1, A, True, Worksheet, Null, Empty)", 2)

    actual = ArrayUtils.CStrArray("A", Array("B", Array("C", Array("D", emptyArr, "E"))))
    Call PrintResult(LangUtils.ToString(actual) = "(A, B, C, D, E)", 3)

End Sub

Private Sub TestConcat()
    Dim emptyArr()      As Variant
    Dim arr1()          As String
    Dim arr2()          As Integer
    Dim arr3()          As Variant
    Dim arr4()          As Variant

    Debug.Print "--- TestConcat ---"

    ReDim arr1(2)
    arr1(0) = "A"
    arr1(1) = "B"
    arr1(2) = "C"
    ReDim arr2(2)
    arr2(0) = 1
    arr2(1) = 2
    arr2(2) = 3
    arr3 = Array(New MyClass, New MyClass)
    arr4 = Array(New MyClass, New MyClass)

    Call ArrayUtils.Concat(emptyArr, emptyArr)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    Call ArrayUtils.Concat(arr1, emptyArr)
    Call PrintResult(LangUtils.ToString(arr1) = "(A, B, C)", 2)

    Call ArrayUtils.Concat(emptyArr, arr2)
    Call PrintResult(LangUtils.ToString(emptyArr) = "(1, 2, 3)", 3)

    ' 暗黙の型変換可能なら異なるデータ型配列でもOK
    Call ArrayUtils.Concat(arr1, arr2)
    Call PrintResult(LangUtils.ToString(arr1) = "(A, B, C, 1, 2, 3)", 4)

    ' 可変長引数の確認
    Call ArrayUtils.Concat(arr1)
    Call PrintResult(LangUtils.ToString(arr1) = "(A, B, C, 1, 2, 3)", 5)
    Call ArrayUtils.Concat(arr1, "D", "E", Array("F"))
    Call PrintResult(LangUtils.ToString(arr1) = "(A, B, C, 1, 2, 3, D, E, F)", 6)

    ' 暗黙の型変換不可なら型不一致エラー
    On Error Resume Next
    Call ArrayUtils.Concat(arr2, arr1)
    Call PrintResult(Err.Number = 13, 7)
    On Error GoTo 0

    ' オブジェクト型でも確認
    Call ArrayUtils.Concat(arr3, arr4)
    Call PrintResult(Length(arr3) = 4, 8)

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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(Not ArrayUtils.ContainsAll(emptyArr, items1), 2)
    Call PrintResult(ArrayUtils.ContainsAll(arr, emptyArr), 3)

    ' 要素が1つも見つからない場合
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items1), 4)
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items2), 5)

    ' 一部の要素が見つかる場合
    items2(0) = 1
    items2(1) = "A"
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, items2), 6)

    ' すべての要素が見つかる場合
    items1(0) = 1
    items2(2) = "C"
    Call PrintResult(ArrayUtils.ContainsAll(arr, items1), 7)
    Call PrintResult(ArrayUtils.ContainsAll(arr, items2), 8)

    ' 可変長引数の確認
    Call PrintResult(ArrayUtils.ContainsAll(arr), 9)
    Call PrintResult(ArrayUtils.ContainsAll(arr, 1, 2, 3), 10)
    Call PrintResult(Not ArrayUtils.ContainsAll(arr, 1, 2, Array(4)), 11)

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
    Call PrintResult(Err.Number = 5, 1)
    On Error GoTo 0

    ' 空配列の場合
    Call PrintResult(Not ArrayUtils.ContainsAny(emptyArr, items1), 2)
    Call PrintResult(ArrayUtils.ContainsAny(arr, emptyArr), 3)

    ' 要素が1つも見つからない場合
    Call PrintResult(Not ArrayUtils.ContainsAny(arr, items1), 4)
    Call PrintResult(Not ArrayUtils.ContainsAny(arr, items2), 5)

    ' 一部の要素が見つかる場合
    items2(1) = "A"
    Call PrintResult(ArrayUtils.ContainsAny(arr, items2), 6)

    ' すべての要素が見つかる場合
    items1(0) = 1
    items2(0) = 1
    items2(2) = "C"
    Call PrintResult(ArrayUtils.ContainsAny(arr, items1), 7)
    Call PrintResult(ArrayUtils.ContainsAny(arr, items2), 8)

    ' 可変長引数の確認
    Call PrintResult(ArrayUtils.ContainsAny(arr), 9)
    Call PrintResult(ArrayUtils.ContainsAny(arr, 1, 2, 3), 10)
    Call PrintResult(ArrayUtils.ContainsAny(arr, 3, 4, Array(5)), 11)
    Call PrintResult(Not ArrayUtils.ContainsAny(arr, Array(4), Array(5, 6)), 12)

End Sub

Private Sub TestCount()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim obj             As Variant

    Debug.Print "--- TestCount ---"

    ' 空配列の場合
    Call PrintResult(ArrayUtils.Count(emptyArr, 1) = 0, 1)

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 2
    arr(5) = 3
    arr(6) = 3

    Call PrintResult(ArrayUtils.Count(arr, 0) = 0, 2)
    Call PrintResult(ArrayUtils.Count(arr, 1) = 1, 3)
    Call PrintResult(ArrayUtils.Count(arr, 3) = 3, 4)

    arr(1) = 0
    arr(2) = 0
    arr(3) = 0
    arr(4) = 0
    arr(5) = 0
    arr(6) = 0

    Call PrintResult(ArrayUtils.Count(arr, 0) = 6, 5)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Set arr(1) = obj
    Set arr(2) = New MyClass
    Set arr(3) = New MyClass
    Set arr(4) = obj
    Set arr(5) = New MyClass
    Set arr(6) = obj

    Call PrintResult(ArrayUtils.Count(arr, obj) = 3, 6)

End Sub

Private Sub TestFill()
    Dim emptyArr()      As Variant
    Dim arr1(1 To 1)    As Variant
    Dim arr2(1 To 6)    As Variant
    Dim obj             As Variant

    Debug.Print "--- TestFill ---"

    ' 空配列の場合
    Call ArrayUtils.Fill(emptyArr, 1)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    arr1(1) = 1

    Call ArrayUtils.Fill(arr1, 1)
    Call PrintResult(LangUtils.ToString(arr1) = "(1)", 2)
    Call ArrayUtils.Fill(arr1, 2, 1)
    Call PrintResult(LangUtils.ToString(arr1) = "(2)", 3)
    Call ArrayUtils.Fill(arr1, 3, 1, 1)
    Call PrintResult(LangUtils.ToString(arr1) = "(3)", 4)

    arr2(1) = 0
    arr2(2) = 0
    arr2(3) = 0
    arr2(4) = 0
    arr2(5) = 0
    arr2(6) = 0

    ' 配列インデックス範囲外の場合
    On Error Resume Next
    Call ArrayUtils.Fill(arr2, 1, 7)
    Call PrintResult(Err.Number = 9, "5-1")
    Call PrintResult(LangUtils.ToString(arr2) = "(0, 0, 0, 0, 0, 0)", "5-2")
    On Error GoTo 0

    Call ArrayUtils.Fill(arr2, 1)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 1, 1, 1, 1, 1)", 6)
    Call ArrayUtils.Fill(arr2, 2, 1)
    Call PrintResult(LangUtils.ToString(arr2) = "(2, 2, 2, 2, 2, 2)", 7)
    Call ArrayUtils.Fill(arr2, 3, 4)
    Call PrintResult(LangUtils.ToString(arr2) = "(2, 2, 2, 3, 3, 3)", 8)
    Call ArrayUtils.Fill(arr2, 4, 6)
    Call PrintResult(LangUtils.ToString(arr2) = "(2, 2, 2, 3, 3, 4)", 9)
    Call ArrayUtils.Fill(arr2, 5, 1, 4)
    Call PrintResult(LangUtils.ToString(arr2) = "(5, 5, 5, 5, 3, 4)", 10)
    Call ArrayUtils.Fill(arr2, 6, 1, 6)
    Call PrintResult(LangUtils.ToString(arr2) = "(6, 6, 6, 6, 6, 6)", 11)
    Call ArrayUtils.Fill(arr2, 7, 3, 2)
    Call PrintResult(LangUtils.ToString(arr2) = "(6, 6, 7, 7, 6, 6)", 12)
    Call ArrayUtils.Fill(arr2, 8, 6, 1)
    Call PrintResult(LangUtils.ToString(arr2) = "(6, 6, 7, 7, 6, 8)", 13)
    Call ArrayUtils.Fill(arr2, 9, -1)
    Call PrintResult(LangUtils.ToString(arr2) = "(6, 6, 7, 7, 6, 9)", 14)
    Call ArrayUtils.Fill(arr2, 10, -6, 7)
    Call PrintResult(LangUtils.ToString(arr2) = "(10, 10, 10, 10, 10, 10)", 15)

    ' オブジェクト型でも確認
    Set obj = New MyClass

    Call ArrayUtils.Fill(arr2, obj)
    Call PrintResult(arr2(1) Is obj, "16-1")
    Call PrintResult(arr2(6) Is obj, "16-2")

End Sub

Private Sub TestFirst()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestFirst ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.First(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr = Array(1, 2, 3)

    Call PrintResult(ArrayUtils.First(arr) = 1, 2)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(obj, New MyClass, New MyClass)

    Call PrintResult(ArrayUtils.First(arr) Is obj, 3)

End Sub

Private Sub TestFlatten()
    Dim emptyArr()      As Variant
    Dim intArr(2)       As Integer
    Dim strArr(2)       As String
    Dim arr             As Variant
    Dim obj             As Variant
    Dim actual()        As Variant

    Debug.Print "--- TestFlatten ---"

    actual = ArrayUtils.Flatten(0)
    Call PrintResult(LangUtils.ToString(actual) = "()", 1)

    actual = ArrayUtils.Flatten(0, emptyArr)
    Call PrintResult(LangUtils.ToString(actual) = "()", 2)

    actual = ArrayUtils.Flatten(0, 1)
    Call PrintResult(LangUtils.ToString(actual) = "(1)", 3)

    actual = ArrayUtils.Flatten(0, 1, 2, "A")
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, A)", 4)

    actual = ArrayUtils.Flatten(0, Array("A", "B", "C"))
    Call PrintResult(LangUtils.ToString(actual) = "(A, B, C)", 5)

    intArr(0) = 1
    intArr(1) = 2
    intArr(2) = 3
    strArr(0) = "A"
    strArr(1) = "B"
    strArr(2) = "C"
    actual = ArrayUtils.Flatten(0, intArr, emptyArr, strArr, True)
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, A, B, C, True)", 6)

    arr = Array(intArr, strArr)

    actual = ArrayUtils.Flatten(1, Array(arr, "a", "b"))
    Call PrintResult(LangUtils.ToString(actual) = "(((1, 2, 3), (A, B, C)), a, b)", 7)

    actual = ArrayUtils.Flatten(2, Array(arr, "a", "b"))
    Call PrintResult(LangUtils.ToString(actual) = "((1, 2, 3), (A, B, C), a, b)", 8)

    actual = ArrayUtils.Flatten(3, Array(arr, "a", "b"))
    Call PrintResult(LangUtils.ToString(actual) = "(1, 2, 3, A, B, C, a, b)", 9)

    ' オブジェクト型でも確認
    actual = ArrayUtils.Flatten(0, New MyClass, Array(New MyClass, Array(New MyClass)))
    Call PrintResult(ArrayUtils.Length(actual) = 3, 10)

End Sub

Private Sub TestGetAt()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestGetAt ---"

    ' 空配列の場合
    Call PrintResult(IsNull(ArrayUtils.GetAt(emptyArr, 0)), 1)

    arr = Array(1, 2, 3)

    Call PrintResult(ArrayUtils.GetAt(arr, 0) = 1, 2)
    Call PrintResult(ArrayUtils.GetAt(arr, 1) = 2, 3)
    Call PrintResult(ArrayUtils.GetAt(arr, 2) = 3, 4)
    Call PrintResult(IsNull(ArrayUtils.GetAt(arr, 3)), 5)
    Call PrintResult(ArrayUtils.GetAt(arr, -1) = 3, 6)
    Call PrintResult(ArrayUtils.GetAt(arr, -3) = 1, 7)
    Call PrintResult(ArrayUtils.GetAt(arr, -4, 0) = 0, 8)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(obj, New MyClass, New MyClass)

    Call PrintResult(ArrayUtils.GetAt(arr, 0) Is obj, 9)
    Call PrintResult(ArrayUtils.GetAt(arr, 3, Nothing) Is Nothing, 10)

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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
    On Error GoTo 0

    On Error Resume Next
    Call ArrayUtils.IsEqual(arr1, 123)
    Call PrintResult(Err.Number = 5, 2)
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

Private Sub TestJoin(ByVal arr As Variant)
    Dim emptyArr()      As Variant
    Dim obj             As Variant

    Debug.Print "--- TestJoin ---"

    ' 配列以外の場合
    Call ArrayUtils.Join(emptyArr, ",")
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    Debug.Print ArrayUtils.Join(arr)
    Debug.Print ArrayUtils.Join(arr, " | ")

End Sub

Private Sub TestLast()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestLast ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Last(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr = Array(1, 2, 3)

    Call PrintResult(ArrayUtils.Last(arr) = 3, 2)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(New MyClass, New MyClass, obj)

    Call PrintResult(ArrayUtils.Last(arr) Is obj, 3)

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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
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
    Call PrintResult(Err.Number = 5, 1)
    On Error GoTo 0

    ' 空配列
    Call PrintResult(ArrayUtils.Length(testArr) = 0, 2)

    ' 空配列以外
    Call PrintResult(ArrayUtils.Length(testArr2) = 1, 3)
    Call PrintResult(ArrayUtils.Length(arr) = 18, 4)

End Sub

Private Sub TestMax()
    Dim emptyArr()      As Variant

    Debug.Print "--- TestMax ---"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.Max("ABC")
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Max(emptyArr)
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    ' 要素一つ
    Call PrintResult(ArrayUtils.Max(Array(1)) = 1, 3)

    ' 複数要素
    Call PrintResult(ArrayUtils.Max(Array(1, 2, 3, 2, 3, 3)) = 3, 4)
    Call PrintResult(ArrayUtils.Max(Array("A", "B", "C")) = "C", 5)

End Sub

Private Sub TestMin()
    Dim emptyArr()      As Variant

    Debug.Print "--- TestMin ---"

    ' 配列以外の場合
    On Error Resume Next
    Call ArrayUtils.Min("ABC")
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Min(emptyArr)
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    ' 要素一つ
    Call PrintResult(ArrayUtils.Min(Array(1)) = 1, 3)

    ' 複数要素
    Call PrintResult(ArrayUtils.Min(Array(1, 2, 3, 2, 3, 3)) = 1, 4)
    Call PrintResult(ArrayUtils.Min(Array("A", "B", "C")) = "A", 5)

End Sub

Private Sub TestPop()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim removed         As Variant
    Dim obj             As Variant

    Debug.Print "--- TestPop ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Pop(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr = Array(1, 2, 3)

    removed = ArrayUtils.Pop(arr)
    Call PrintResult(removed = 3, 2)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2)", 3)

    removed = ArrayUtils.Pop(arr)
    Call PrintResult(removed = 2, 4)
    Call PrintResult(LangUtils.ToString(arr) = "(1)", 5)

    removed = ArrayUtils.Pop(arr)
    Call PrintResult(removed = 1, 6)
    Call PrintResult(LangUtils.ToString(arr) = "()", 7)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(New MyClass, New MyClass, obj)

    Set removed = ArrayUtils.Pop(arr)
    Call PrintResult(removed Is obj, 8)
    Call PrintResult(ArrayUtils.Length(arr) = 2, 9)

End Sub

Private Sub TestPush()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestPush ---"

    arr = Array(1, 2, 3)

    ' 空配列の場合
    Call ArrayUtils.Push(emptyArr, "A")
    Call PrintResult(LangUtils.ToString(emptyArr) = "(A)", 1)

    Call ArrayUtils.Push(arr, "A")
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, A)", 2)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Call ArrayUtils.Push(arr, obj)
    Call PrintResult(arr(4) Is obj, 3)

End Sub

Private Sub TestRemove()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestRemove ---"

    arr = Array(1, 2, 3, 2, 3, 3)

    ' 空配列の場合
    Call ArrayUtils.Remove(emptyArr, 1)
    Call PrintResult(IsEmptyArray(emptyArr), 1)

    ' (1, 2, 3, 2, 3, 3) => (2, 3, 2, 3, 3)
    Call ArrayUtils.Remove(arr, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 3, 2, 3, 3)", 2)

    ' (2, 3, 2, 3, 3) => (2, 2)
    Call ArrayUtils.Remove(arr, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 2)", 3)

    ' (2, 2) => (2)
    Call ArrayUtils.Remove(arr, 2, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(2)", 4)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(obj, obj, obj, obj, obj, obj)

    Call ArrayUtils.Remove(arr, obj, 3)
    Call PrintResult(Length(arr) = 3, 5)

    Call ArrayUtils.Remove(arr, obj)
    Call PrintResult(IsEmptyArray(arr), 6)

End Sub

Private Sub TestRemoveAll()
    Dim emptyArr()      As Variant
    Dim arr1()          As Variant
    Dim arr2()          As Variant
    Dim obj             As Variant

    Debug.Print "--- TestRemoveAll ---"

    arr1 = Array(1, 2, 3, 2, 3, 3)

    ' 空配列の場合
    Call ArrayUtils.RemoveAll(emptyArr, arr1)
    Call PrintResult(IsEmptyArray(emptyArr), 1)

    ' 削除する要素なし
    arr2 = Array("A", "B")
    Call ArrayUtils.RemoveAll(arr1, arr2)
    Call PrintResult(LangUtils.ToString(arr1) = "(1, 2, 3, 2, 3, 3)", 2)

    arr2 = Array(1, 2)
    Call ArrayUtils.RemoveAll(arr1, arr2)
    Call PrintResult(LangUtils.ToString(arr1) = "(3, 3, 3)", 3)

    ' 全削除
    arr1 = Array(1, 2, 3, 2, 3, 3)
    arr2 = Array(1, 2, 3)
    Call ArrayUtils.RemoveAll(arr1, arr2)
    Call PrintResult(LangUtils.ToString(arr1) = "()", 4)

    ' 見つかる要素、見つからない要素混合
    arr1 = Array(1, 2, 3, 2, 3, 3)
    arr2 = Array("A", 3)
    Call ArrayUtils.RemoveAll(arr1, arr2)
    Call PrintResult(LangUtils.ToString(arr1) = "(1, 2, 2)", 5)

    ' 可変長引数の確認
    arr1 = Array(1, 2, 3, 2, 3, 3)
    Call ArrayUtils.RemoveAll(arr1)
    Call PrintResult(LangUtils.ToString(arr1) = "(1, 2, 3, 2, 3, 3)", 6)
    Call ArrayUtils.RemoveAll(arr1, 1, Array(3))
    Call PrintResult(LangUtils.ToString(arr1) = "(2, 2)", 7)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr1 = Array(obj, obj, obj, obj)
    arr2 = Array(obj)

    Call ArrayUtils.RemoveAll(arr1, arr2)
    Call PrintResult(IsEmptyArray(arr1), 8)

End Sub

Private Sub TestRemoveAt()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant
    Dim removed         As Variant

    Debug.Print "--- TestRemoveAt ---"

    arr = Array(1, 2, 3, 4)

    ' 配列インデックス範囲外の場合
    On Error Resume Next
    Call ArrayUtils.RemoveAt(emptyArr, 0)
    Call PrintResult(Err.Number = 9, 1)
    On Error GoTo 0

    On Error Resume Next
    Call ArrayUtils.RemoveAt(arr, 4)
    Call PrintResult(Err.Number = 9, 2)
    On Error GoTo 0

    ' 先頭から削除：(1, 2, 3, 4) => (2, 3, 4)
    removed = ArrayUtils.RemoveAt(arr, 0)
    Call PrintResult(removed = 1, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 3, 4)", 4)

    ' 途中から削除：(2, 3, 4) => (2, 4)
    removed = ArrayUtils.RemoveAt(arr, 1)
    Call PrintResult(removed = 3, 5)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 4)", 6)

    ' 末尾から削除：(2, 4) => (2)
    removed = ArrayUtils.RemoveAt(arr, 1)
    Call PrintResult(removed = 4, 7)
    Call PrintResult(LangUtils.ToString(arr) = "(2)", 8)

    ' 削除した結果空：(2) => ()
    removed = ArrayUtils.RemoveAt(arr, 0)
    Call PrintResult(removed = 2, 9)
    Call PrintResult(IsEmptyArray(arr), 10)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(1, 2, obj, 4)
    Set removed = ArrayUtils.RemoveAt(arr, 2)
    Call PrintResult(removed Is obj, 11)

End Sub

Private Sub TestReplace()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim obj1            As Variant
    Dim obj2            As Variant

    Debug.Print "--- TestReplace ---"

    ' 空配列の場合
    Call PrintResult(ArrayUtils.Replace(emptyArr, 1, "A") = 0, "1-1")
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", "1-2")

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 2
    arr(5) = 3
    arr(6) = 3

    Call PrintResult(ArrayUtils.Replace(arr, 1, "A") = 1, "2-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 2, 3, 2, 3, 3)", "2-2")

    Call PrintResult(ArrayUtils.Replace(arr, 2, "B") = 2, "3-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, 3, B, 3, 3)", "3-2")

    Call PrintResult(ArrayUtils.Replace(arr, 3, "C", 1) = 1, "4-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C, B, 3, 3)", "4-2")

    arr(1) = 1
    arr(2) = 1
    arr(3) = 1
    arr(4) = 1
    arr(5) = 1
    arr(6) = 1

    Call PrintResult(ArrayUtils.Replace(arr, 1, "A") = 6, "5-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, A, A, A, A, A)", "5-2")

    Call PrintResult(ArrayUtils.Replace(arr, "A", "B", 7) = 6, "6-1")
    Call PrintResult(LangUtils.ToString(arr) = "(B, B, B, B, B, B)", "6-2")

    ' オブジェクト型でも確認
    Set obj1 = New MyClass
    Set obj2 = New MyClass
    Set arr(1) = obj1
    Set arr(2) = obj1
    Set arr(3) = obj1

    Call PrintResult(ArrayUtils.Replace(arr, obj1, obj2) = 3, "7-1")
    Call PrintResult(arr(3) Is obj2, "7-2")

End Sub

Private Sub TestReplaceLast()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim obj1            As Variant
    Dim obj2            As Variant

    Debug.Print "--- TestReplaceLast ---"

    ' 空配列の場合
    Call PrintResult(ArrayUtils.ReplaceLast(emptyArr, 1, "A") = 0, "1-1")
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", "1-2")

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 2
    arr(5) = 3
    arr(6) = 3

    Call PrintResult(ArrayUtils.ReplaceLast(arr, 1, "A") = 1, "2-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 2, 3, 2, 3, 3)", "2-2")

    Call PrintResult(ArrayUtils.ReplaceLast(arr, 2, "B") = 2, "3-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, 3, B, 3, 3)", "3-2")

    Call PrintResult(ArrayUtils.ReplaceLast(arr, 3, "C", 1) = 1, "4-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, 3, B, 3, C)", "4-2")

    arr(1) = 1
    arr(2) = 1
    arr(3) = 1
    arr(4) = 1
    arr(5) = 1
    arr(6) = 1

    Call PrintResult(ArrayUtils.ReplaceLast(arr, 1, "A") = 6, "5-1")
    Call PrintResult(LangUtils.ToString(arr) = "(A, A, A, A, A, A)", "5-2")

    Call PrintResult(ArrayUtils.ReplaceLast(arr, "A", "B", 7) = 6, "6-1")
    Call PrintResult(LangUtils.ToString(arr) = "(B, B, B, B, B, B)", "6-2")

    ' オブジェクト型でも確認
    Set obj1 = New MyClass
    Set obj2 = New MyClass
    Set arr(1) = obj1
    Set arr(2) = obj1
    Set arr(3) = obj1

    Call PrintResult(ArrayUtils.ReplaceLast(arr, obj1, obj2) = 3, "7-1")
    Call PrintResult(arr(1) Is obj2, "7-2")

End Sub

Private Sub TestReverse()
    Dim emptyArr()      As Variant
    Dim arr1(1 To 1)    As Variant
    Dim arr2(1 To 3)    As Variant
    Dim arr3(1 To 6)    As Variant
    Dim obj             As Variant

    Debug.Print "--- TestReverse ---"

    ' 空配列の場合
    Call ArrayUtils.Reverse(emptyArr)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    arr1(1) = 1

    Call ArrayUtils.Reverse(arr1)
    Call PrintResult(LangUtils.ToString(arr1) = "(1)", 2)

    arr2(1) = 1
    arr2(2) = 2
    arr2(3) = 3

    Call ArrayUtils.Reverse(arr2)
    Call PrintResult(LangUtils.ToString(arr2) = "(3, 2, 1)", 3)

    arr3(1) = 1
    arr3(2) = 2
    arr3(3) = 3
    arr3(4) = 4
    arr3(5) = 5
    arr3(6) = 6

    Call ArrayUtils.Reverse(arr3)
    Call PrintResult(LangUtils.ToString(arr3) = "(6, 5, 4, 3, 2, 1)", 4)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Set arr3(1) = obj

    Call ArrayUtils.Reverse(arr3)
    Call PrintResult(arr3(6) Is obj, 5)

End Sub

Private Sub TestRotate()
    Dim emptyArr()      As Variant
    Dim arr1(1 To 1)    As Variant
    Dim arr2(1 To 6)    As Variant
    Dim obj             As Variant

    Debug.Print "--- TestRotate ---"

    ' 空配列の場合
    Call ArrayUtils.Rotate(emptyArr, 1)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    arr1(1) = 1

    ' 要素が一つの場合
    Call ArrayUtils.Rotate(arr1, 0)
    Call PrintResult(LangUtils.ToString(arr1) = "(1)", 2)
    Call ArrayUtils.Rotate(arr1, 1)
    Call PrintResult(LangUtils.ToString(arr1) = "(1)", 3)
    Call ArrayUtils.Rotate(arr1, -1)
    Call PrintResult(LangUtils.ToString(arr1) = "(1)", 4)

    arr2(1) = 1
    arr2(2) = 2
    arr2(3) = 3
    arr2(4) = 4
    arr2(5) = 5
    arr2(6) = 6

    ' 位置が変わらない場合
    Call ArrayUtils.Rotate(arr2, 0)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 5)
    Call ArrayUtils.Rotate(arr2, 6)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 6)
    Call ArrayUtils.Rotate(arr2, -12)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 7)

    ' 位置が変わる場合
    Call ArrayUtils.Rotate(arr2, 1)
    Call PrintResult(LangUtils.ToString(arr2) = "(6, 1, 2, 3, 4, 5)", 8)
    Call ArrayUtils.Rotate(arr2, -1)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 9)
    Call ArrayUtils.Rotate(arr2, 3)
    Call PrintResult(LangUtils.ToString(arr2) = "(4, 5, 6, 1, 2, 3)", 10)
    Call ArrayUtils.Rotate(arr2, -3)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 11)
    Call ArrayUtils.Rotate(arr2, 10)
    Call PrintResult(LangUtils.ToString(arr2) = "(3, 4, 5, 6, 1, 2)", 12)
    Call ArrayUtils.Rotate(arr2, -10)
    Call PrintResult(LangUtils.ToString(arr2) = "(1, 2, 3, 4, 5, 6)", 13)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Set arr2(1) = obj

    Call ArrayUtils.Rotate(arr2, 5)
    Call PrintResult(arr2(6) Is obj, 14)

End Sub

Private Sub TestSample()
    Dim i               As Long
    Dim emptyArr()      As Variant
    Dim arr1(1 To 1)    As Variant
    Dim arr2(1 To 7)    As Variant
    Dim result()        As Variant

    Debug.Print "--- TestSample ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Sample(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr1(1) = 1

    result = ArrayUtils.Sample(arr1)
    Debug.Print LangUtils.ToString(result)

    result = ArrayUtils.Sample(arr1, 1)
    Debug.Print LangUtils.ToString(result)

    result = ArrayUtils.Sample(arr1, 2)
    Debug.Print LangUtils.ToString(result)

    arr2(1) = 1
    arr2(2) = 2
    arr2(3) = 3
    arr2(4) = 4
    arr2(5) = 5
    arr2(6) = 6
    Set arr2(7) = New MyClass

    result = ArrayUtils.Sample(arr2)

    Debug.Print "ランダムに一つ取得"
    For i = 1 To 5
        result = ArrayUtils.Sample(arr2)
        Debug.Print "  " & LangUtils.ToString(result)
    Next

    Debug.Print "ランダムに複数取得(重複なし)"
    For i = 1 To 5
        result = ArrayUtils.Sample(arr2, 3)
        Debug.Print "  " & LangUtils.ToString(result)
    Next
    For i = 1 To 5
        result = ArrayUtils.Sample(arr2, 7)
        Debug.Print "  " & LangUtils.ToString(result)
    Next
    result = ArrayUtils.Sample(arr2, 8)
    Debug.Print "  " & LangUtils.ToString(result)

    Debug.Print "ランダムに複数取得(重複あり)"
    For i = 1 To 5
        result = ArrayUtils.Sample(arr2, 3, False)
        Debug.Print "  " & LangUtils.ToString(result)
    Next
    For i = 1 To 5
        result = ArrayUtils.Sample(arr2, 7, False)
        Debug.Print "  " & LangUtils.ToString(result)
    Next

End Sub

Private Sub TestSetAt()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim before          As Variant
    Dim obj1            As Variant
    Dim obj2            As Variant

    Debug.Print "--- TestSetAt ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.SetAt(emptyArr, 0, "A")
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr = Array(1, 2, 3)

    ' 配列インデックス範囲外の場合
    On Error Resume Next
    Call ArrayUtils.SetAt(arr, "A", 3)
    Call PrintResult(Err.Number <> 0, 2)
    On Error GoTo 0

    before = ArrayUtils.SetAt(arr, 0, "A")
    Call PrintResult(before = 1, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(A, 2, 3)", 4)

    before = ArrayUtils.SetAt(arr, 1, "B")
    Call PrintResult(before = 2, 5)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, 3)", 6)

    before = ArrayUtils.SetAt(arr, 2, "C")
    Call PrintResult(before = 3, 7)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C)", 8)

    before = ArrayUtils.SetAt(arr, -1, "a")
    Call PrintResult(before = "C", 9)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, a)", 10)

    before = ArrayUtils.SetAt(arr, -3, "c")
    Call PrintResult(before = "A", 11)
    Call PrintResult(LangUtils.ToString(arr) = "(c, B, a)", 12)

    ' オブジェクト型でも確認
    Set obj1 = New MyClass
    Set obj2 = New MyClass
    arr = Array(obj1, New MyClass, New MyClass)

    Set before = ArrayUtils.SetAt(arr, 0, obj2)
    Call PrintResult(before Is obj1, 13)
    Call PrintResult(arr(0) Is obj2, 14)

End Sub

Private Sub TestShift()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim removed         As Variant
    Dim obj             As Variant

    Debug.Print "--- TestShift ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Shift(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr = Array(1, 2, 3)

    removed = ArrayUtils.Shift(arr)
    Call PrintResult(removed = 1, 2)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 3)", 3)

    removed = ArrayUtils.Shift(arr)
    Call PrintResult(removed = 2, 4)
    Call PrintResult(LangUtils.ToString(arr) = "(3)", 5)

    removed = ArrayUtils.Shift(arr)
    Call PrintResult(removed = 3, 6)
    Call PrintResult(LangUtils.ToString(arr) = "()", 7)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(obj, New MyClass, New MyClass)

    Set removed = ArrayUtils.Shift(arr)
    Call PrintResult(removed Is obj, 8)
    Call PrintResult(Length(arr) = 2, 9)

End Sub

Private Sub TestShuffle()
    Dim i               As Long
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant

    Debug.Print "--- TestShuffle ---"

    ' 空配列の場合
    Call ArrayUtils.Shuffle(emptyArr)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 4
    arr(5) = 5
    Set arr(6) = New MyClass

    For i = 1 To 10
        Call ArrayUtils.Shuffle(arr)
        Debug.Print LangUtils.ToString(arr)
    Next

End Sub

Private Sub TestSlice()
    Dim i               As Long
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim result()        As Variant
    Dim obj             As Variant

    Debug.Print "--- TestSlice ---"

    ' 空配列の場合
    On Error Resume Next
    result = ArrayUtils.Slice(emptyArr)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 4
    arr(5) = 5
    arr(6) = 6

    result = ArrayUtils.Slice(arr)
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3, 4, 5, 6)", 2)

    result = ArrayUtils.Slice(arr, 1)
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3, 4, 5, 6)", 3)

    result = ArrayUtils.Slice(arr, 4)
    Call PrintResult(LangUtils.ToString(result) = "(4, 5, 6)", 4)

    result = ArrayUtils.Slice(arr, 6)
    Call PrintResult(LangUtils.ToString(result) = "(6)", 5)

    result = ArrayUtils.Slice(arr, 1, 7)
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3, 4, 5, 6)", 6)

    result = ArrayUtils.Slice(arr, 1, 4)
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3)", 7)

    result = ArrayUtils.Slice(arr, 5, 6)
    Call PrintResult(LangUtils.ToString(result) = "(5)", 8)

    result = ArrayUtils.Slice(arr, 3, 3)
    Call PrintResult(LangUtils.ToString(result) = "()", 9)

    result = ArrayUtils.Slice(arr, 7, 7)
    Call PrintResult(LangUtils.ToString(result) = "()", 10)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Set arr(1) = obj

    result = ArrayUtils.Slice(arr, 1, 2)
    Call PrintResult(result(0) Is obj, 11)

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
    Call PrintResult(Err.Number = 5, 1)
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

Private Sub TestSplice()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim added()         As Variant
    Dim obj             As Variant

    Debug.Print "--- TestSplice ---"

    added = Array("A", "B", "C")

    ' 空配列の場合
    Call ArrayUtils.Splice(emptyArr, 0)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    Call ArrayUtils.Splice(emptyArr, 0, 0, emptyArr)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 2)

    Call ArrayUtils.Splice(emptyArr, 0, 0, added)
    Call PrintResult(LangUtils.ToString(emptyArr) = "(A, B, C)", 3)

    On Error Resume Next
    Call ArrayUtils.Splice(emptyArr, 10, 0, added)
    Call PrintResult(Err.Number = 9, 4)
    On Error GoTo 0

    ' 削除のみ
    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0)
    Call PrintResult(LangUtils.ToString(arr) = "()", 5)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 6)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 5)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5)", 7)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 6)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5, 6)", 8)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(2, 3, 4, 5, 6)", 9)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(4, 5, 6)", 10)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 6)
    Call PrintResult(LangUtils.ToString(arr) = "()", 11)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 7)
    Call PrintResult(LangUtils.ToString(arr) = "()", 12)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 4, 5, 6)", 13)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 6)", 14)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 4)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2)", 15)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 5)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2)", 16)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 5, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5)", 17)

    emptyArr = Array()
    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 5, 1, emptyArr)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5)", 18)

    ' 追加のみ
    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 0, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C, 1, 2, 3, 4, 5, 6)", 19)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 0, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, A, B, C, 3, 4, 5, 6)", 20)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 5, 0, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5, A, B, C, 6)", 21)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 6, 0, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5, 6, A, B, C)", 22)

    ' 削除しつつ追加
    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 1, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C, 2, 3, 4, 5, 6)", 23)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 3, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C, 4, 5, 6)", 24)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 5, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C, 6)", 25)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 6, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C)", 26)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 0, 7, added)
    Call PrintResult(LangUtils.ToString(arr) = "(A, B, C)", 27)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 1, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, A, B, C, 4, 5, 6)", 28)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 3, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, A, B, C, 6)", 29)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 4, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, A, B, C)", 30)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 5, 1, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5, A, B, C)", 31)

    arr = Array(1, 2, 3, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 6, 1, added)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3, 4, 5, 6, A, B, C)", 32)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    arr = Array(1, 2, New MyClass, 4, 5, 6)
    Call ArrayUtils.Splice(arr, 2, 1, Array(obj))
    Call PrintResult(arr(2) Is obj, 33)

End Sub

Private Sub TestSwap()
    Dim emptyArr()      As Variant
    Dim arr(1 To 3)     As Variant
    Dim obj             As Variant

    Debug.Print "--- TestSwap ---"

    ' 空配列の場合
    On Error Resume Next
    Call ArrayUtils.Swap(emptyArr, 0, 1)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    arr(1) = 1
    arr(2) = 2
    arr(3) = 3

    Call ArrayUtils.Swap(arr, 1, 1)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 2)

    Call ArrayUtils.Swap(arr, 1, 3)
    Call PrintResult(LangUtils.ToString(arr) = "(3, 2, 1)", 3)

    Call ArrayUtils.Swap(arr, 3, 2)
    Call PrintResult(LangUtils.ToString(arr) = "(3, 1, 2)", 4)

    Set obj = New MyClass
    Set arr(1) = obj
    Set arr(2) = New MyClass
    Set arr(3) = New MyClass

    ' オブジェクト型でも確認
    Call ArrayUtils.Swap(arr, 1, 2)
    Call PrintResult(arr(2) Is obj, 5)

End Sub

Private Sub TestUnique()
    Dim emptyArr()      As Variant
    Dim arr             As Variant
    Dim obj             As Variant

    Debug.Print "--- TestUnique ---"

    ' 空配列の場合
    Call ArrayUtils.Unique(emptyArr)
    Call PrintResult(LangUtils.ToString(emptyArr) = "()", 1)

    arr = Array(1)

    Call ArrayUtils.Unique(arr)
    Call PrintResult(LangUtils.ToString(arr) = "(1)", 2)

    arr = Array(1, 2, 3)

    Call ArrayUtils.Unique(arr)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 3)

    arr = Array(1, 2, 3, 2, 3, 3)

    Call ArrayUtils.Unique(arr)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 4)

    arr = Array(4, 1, 1, 4, 1, 5)

    Call ArrayUtils.Unique(arr)
    Call PrintResult(LangUtils.ToString(arr) = "(4, 1, 5)", 5)

    arr = Array(1, 1, 1, 1, 1, 1)

    Call ArrayUtils.Unique(arr)
    Call PrintResult(LangUtils.ToString(arr) = "(1)", 6)

    Set obj = New MyClass
    arr = Array(obj, obj, obj)

    ' オブジェクト型でも確認
    Call ArrayUtils.Unique(arr)
    Call PrintResult(ArrayUtils.Length(arr) = 1, 7)

End Sub

Private Sub TestUnshift()
    Dim emptyArr()      As Variant
    Dim arr()           As Variant
    Dim obj             As Variant

    Debug.Print "--- TestUnshift ---"

    arr = Array(1, 2, 3)

    ' 空配列の場合
    Call ArrayUtils.Unshift(emptyArr, "A")
    Call PrintResult(LangUtils.ToString(emptyArr) = "(A)", 1)

    Call ArrayUtils.Unshift(arr, "A")
    Call PrintResult(LangUtils.ToString(arr) = "(A, 1, 2, 3)", 2)

    ' オブジェクト型でも確認
    Set obj = New MyClass
    Call ArrayUtils.Unshift(arr, obj)
    Call PrintResult(arr(0) Is obj, 3)

End Sub

Private Sub TestValuesAt()
    Dim emptyArr()      As Variant
    Dim arr(1 To 6)     As Variant
    Dim result()        As Variant
    Dim obj             As Variant

    Debug.Print "--- TestValuesAt ---"

    ' 空配列の場合
    On Error Resume Next
    result = ArrayUtils.ValuesAt(emptyArr, 1)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    Set obj = New MyClass
    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    arr(4) = 4
    arr(5) = 5
    Set arr(6) = obj

    result = ArrayUtils.ValuesAt(arr)
    Call PrintResult(LangUtils.ToString(result) = "()", 2)

    result = ArrayUtils.ValuesAt(arr, 1)
    Call PrintResult(LangUtils.ToString(result) = "(1)", 3)

    result = ArrayUtils.ValuesAt(arr, 1, 6)
    Call PrintResult(result(0) = 1, "4-1")
    Call PrintResult(result(1) Is obj, "4-2")

    result = ArrayUtils.ValuesAt(arr, Array(1, 2, 3))
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3)", 5)

    result = ArrayUtils.ValuesAt(arr, 1, 2, Array(3, 4), 5, 1)
    Call PrintResult(LangUtils.ToString(result) = "(1, 2, 3, 4, 5, 1)", 6)

End Sub
