Attribute VB_Name = "TestArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : ArrayUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)

    Debug.Print "=== TestArrayUtils ==="

    Call TestIsEmptyArray(arr)
    Call TestIsEqual(arr)
    Call TestLength(arr)
    Call TestSort

End Sub

Private Sub TestIsEmptyArray(ByVal arr As Variant)
    Dim testArr()       As Variant

    Debug.Print "--- TestIsEmptyArray ---"

    ' 空配列の場合
    Call PrintResult(ArrayUtils.IsEmptyArray(testArr), 1)
    
    ' 空配列以外の以外
    Call PrintResult(Not ArrayUtils.IsEmptyArray(arr), 2)

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

    Call PrintResult(ArrayUtils.IsEqual(emptyArr1, emptyArr2), 1)
    Call PrintResult(Not ArrayUtils.IsEqual(emptyArr1, arr1), 2)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, emptyArr1), 3)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr2), 4)
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr3), 5)
    Call PrintResult(Not ArrayUtils.IsEqual(arr2, arr3), 6)
    Call PrintResult(ArrayUtils.IsEqual(arr1, arr4), 7)
    
    arr4(5) = 6
    
    Call PrintResult(Not ArrayUtils.IsEqual(arr1, arr4), 8)
    Call PrintResult(ArrayUtils.IsEqual(arr, arr), 9)

End Sub

Private Sub TestLength(ByVal arr As Variant)
    Dim testArr()       As Variant
    Dim testArr2(0)     As Variant

    Debug.Print "--- TestLength ---"

    ' 空配列
    Call PrintResult(ArrayUtils.Length(testArr) = 0, 1)

    ' 空配列以外
    Call PrintResult(ArrayUtils.Length(testArr2) = 1, 2)
    Call PrintResult(ArrayUtils.Length(arr) = 18, 3)

End Sub

Private Sub TestSort()
    Dim i               As Long
    Dim arr1(1000)      As Long
    Dim arr2(1000)      As String
    Dim arr3(5)         As MyClass
    Dim result          As Boolean

    Debug.Print "--- TestSort ---"
 
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

    Call PrintResult(result, 1)
 
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

    Call PrintResult(result, 2)

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

    Call PrintResult(result, 3)

End Sub
