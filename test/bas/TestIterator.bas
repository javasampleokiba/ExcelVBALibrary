Attribute VB_Name = "TestIterator"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : Iteratorのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll()

    Debug.Print "=== TestIterator ==="

    Call TestArrayIterator
    Call TestListIterator

End Sub

Private Sub TestArrayIterator()
    Dim it          As Iterator
    Dim arr()       As Integer
    Dim expected    As Integer
    Dim item        As Integer
    Dim count       As Integer

    Debug.Print "--- TestArrayIterator ---"

    Set it = New Iterator

    ' 配列未初期化の場合
    Call it.InitArrayIterator(arr)
    Call PrintResult(Not it.HasNext(), 1)

    arr = ArrayUtils.CIntArray(1)

    ' 要素が一つの場合
    Call it.InitArrayIterator(arr)
    expected = 0
    Do While it.HasNext()
        expected = expected + 1
        Call PrintResult(expected = it.GetNext(), "2-1")
        Call PrintResult(expected - 1 = it.GetCurrentIndex(), "2-2")
        Call PrintResult(1 = it.GetCount(), "2-3")
        Call PrintResult(it.IsFirst(), "2-4")
        Call PrintResult(it.IsLast(), "2-5")
    Loop

    arr = ArrayUtils.CIntArray(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    ' 通常イテレートの確認
    Call it.InitArrayIterator(arr)
    expected = 0
    count = 0
    Do While it.HasNext()
        expected = expected + 1
        count = count + 1
        Call PrintResult(expected = it.GetNext(), "3-1")
        Call PrintResult(expected - 1 = it.GetCurrentIndex(), "3-2")
        Call PrintResult(count = it.GetCount(), "3-3")
        If count = 1 Then
            Call PrintResult(it.IsFirst(), "3-4")
            Call PrintResult(Not it.IsLast(), "3-5")
        ElseIf count = 10 Then
            Call PrintResult(Not it.IsFirst(), "3-6")
            Call PrintResult(it.IsLast(), "3-7")
        Else
            Call PrintResult(Not it.IsFirst(), "3-8")
            Call PrintResult(Not it.IsLast(), "3-9")
        End If
    Loop

    ' 逆順イテレートの確認
    Call it.InitArrayIterator(arr, "Reverse")
    expected = 0
    count = 0
    Do While it.HasNext()
        expected = expected + 1
        count = count + 1
        Call PrintResult(11 - expected = it.GetNext(), "4-1")
        Call PrintResult(11 - expected - 1 = it.GetCurrentIndex(), "4-2")
        Call PrintResult(count = it.GetCount(), "4-3")
        If count = 1 Then
            Call PrintResult(it.IsFirst(), "4-4")
            Call PrintResult(Not it.IsLast(), "4-5")
        ElseIf count = 10 Then
            Call PrintResult(Not it.IsFirst(), "4-6")
            Call PrintResult(it.IsLast(), "4-7")
        Else
            Call PrintResult(Not it.IsFirst(), "4-8")
            Call PrintResult(Not it.IsLast(), "4-9")
        End If
    Loop

    Call it.InitArrayIterator(arr)
    On Error Resume Next
    item = it.GetNext()
    Call it.Remove
    Call PrintResult(Err.Number = 17, 5)
    On Error GoTo 0

End Sub

Private Sub TestListIterator()
    Dim it          As Iterator
    Dim lst         As List
    Dim lst1        As List
    Dim lst2        As List
    Dim expected    As Integer
    Dim item        As Integer
    Dim lstItem     As List
    Dim count       As Integer

    Debug.Print "--- TestListIterator ---"

    Set lst = New List
    Set it = New Iterator

    ' 要素がゼロの場合
    Call it.InitListIterator(lst)
    Call PrintResult(Not it.HasNext(), 1)

    Call lst.Push(1)

    ' 要素が一つの場合
    Call it.InitListIterator(lst)
    expected = 0
    Do While it.HasNext()
        expected = expected + 1
        Call PrintResult(expected = it.GetNext(), "2-1")
        Call PrintResult(expected - 1 = it.GetCurrentIndex(), "2-2")
        Call PrintResult(1 = it.GetCount(), "2-3")
        Call PrintResult(it.IsFirst(), "2-4")
        Call PrintResult(it.IsLast(), "2-5")
    Loop

    Call lst.Concat(2, 3, 4, 5, 6, 7, 8, 9, 10)

    ' 通常イテレートの確認
    Call it.InitListIterator(lst)
    expected = 0
    count = 0
    Do While it.HasNext()
        expected = expected + 1
        count = count + 1
        Call PrintResult(expected = it.GetNext(), "3-1")
        Call PrintResult(expected - 1 = it.GetCurrentIndex(), "3-2")
        Call PrintResult(count = it.GetCount(), "3-3")
        If count = 1 Then
            Call PrintResult(it.IsFirst(), "3-4")
            Call PrintResult(Not it.IsLast(), "3-5")
        ElseIf count = 10 Then
            Call PrintResult(Not it.IsFirst(), "3-6")
            Call PrintResult(it.IsLast(), "3-7")
        Else
            Call PrintResult(Not it.IsFirst(), "3-8")
            Call PrintResult(Not it.IsLast(), "3-9")
        End If
    Loop

    ' 逆順イテレートの確認
    Call it.InitListIterator(lst, "REVERSE")
    expected = 0
    count = 0
    Do While it.HasNext()
        expected = expected + 1
        count = count + 1
        Call PrintResult(11 - expected = it.GetNext(), "4-1")
        Call PrintResult(11 - expected - 1 = it.GetCurrentIndex(), "4-2")
        Call PrintResult(count = it.GetCount(), "4-3")
        If count = 1 Then
            Call PrintResult(it.IsFirst(), "4-4")
            Call PrintResult(Not it.IsLast(), "4-5")
        ElseIf count = 10 Then
            Call PrintResult(Not it.IsFirst(), "4-6")
            Call PrintResult(it.IsLast(), "4-7")
        Else
            Call PrintResult(Not it.IsFirst(), "4-8")
            Call PrintResult(Not it.IsLast(), "4-9")
        End If
    Loop

    ' 削除の確認
    Call it.InitListIterator(lst)
    Do While it.HasNext()
        item = it.GetNext()
        If item = 4 Then
            Call PrintResult(3 = it.GetCurrentIndex(), "5-1")
            Call it.Remove
            Call PrintResult(2 = it.GetCurrentIndex(), "5-2")
        ElseIf item = 6 Or item = 7 Then
            Call it.Remove
        ElseIf item = 10 Then
            Call it.Remove
        End If
    Loop
    Call PrintResult(lst.ToString() = "(1, 2, 3, 5, 8, 9)", "5-3")

    ' 削除の確認（逆順）
    Call it.InitListIterator(lst, "reverse")
    Do While it.HasNext()
        item = it.GetNext()
        If item = 9 Then
            Call PrintResult(5 = it.GetCurrentIndex(), "6-1")
            Call it.Remove
            Call PrintResult(5 = it.GetCurrentIndex(), "6-2")
        ElseIf item = 1 Then
            Call it.Remove
        End If
    Loop
    Call PrintResult(lst.ToString() = "(2, 3, 5, 8)", "6-3")

    ' 全削除の確認
    Call it.InitListIterator(lst)
    count = 0
    Do While it.HasNext()
        count = count + 1
        item = it.GetNext()
        Call it.Remove
        Call PrintResult(-1 = it.GetCurrentIndex(), "7-1")
        Call PrintResult(count = it.GetCount(), "7-2")
        If count = 1 Then
            Call PrintResult(it.IsFirst(), "7-3")
            Call PrintResult(Not it.IsLast(), "7-4")
        ElseIf count = 4 Then
            Call PrintResult(Not it.IsFirst(), "7-5")
            Call PrintResult(it.IsLast(), "7-6")
        Else
            Call PrintResult(Not it.IsFirst(), "7-7")
            Call PrintResult(Not it.IsLast(), "7-8")
        End If
    Loop
    Call PrintResult(lst.ToString() = "()", "7-9")

    Set lst1 = New List
    Call lst1.Concat(1, 2, 3)
    Set lst2 = New List
    Call lst2.Concat("A", "B", "C")
    Call lst.Push(lst1)
    Call lst.Push(lst2)

    ' オブジェクト型でも確認
    Call it.InitListIterator(lst)
    Do While it.HasNext()
        Set lstItem = it.GetNext()
        If it.IsFirst() Then
            Call PrintResult(lstItem.ToString() = "(1, 2, 3)", "8-1")
        Else
            Call it.Remove
        End If
    Loop
    Call PrintResult(lst.ToString() = "((1, 2, 3))", "8-2")

End Sub
