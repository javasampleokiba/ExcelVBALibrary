Attribute VB_Name = "TestList"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : Listのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)
    Dim lst     As List

    Debug.Print "=== TestList ==="

    Set lst = New List
    Call lst.Concat(1, 2, 3, 2, 3, 3)

    Call TestAdd
    Call TestClear(lst.Clone())
    Call TestClone
    Call TestConcat
    Call TestContains(lst.Clone())
    Call TestContainsAll(lst.Clone())
    Call TestContainsAny(lst.Clone())
    Call TestCount(lst.Clone())
    Call TestFirst(lst.Clone())
    Call TestGetAt(lst.Clone())
    Call TestGetDataType(lst.Clone())
    Call TestIndexOf(lst.Clone())
    Call TestIndicesOf(lst.Clone())
    Call TestIsEmptyList(lst.Clone())
    Call TestIsEqual(lst.Clone())
    Call TestJoin
    Call TestLast(lst.Clone())
    Call TestLastIndexOf(lst.Clone())
    Call TestLastIndicesOf(lst.Clone())
    Call TestLength(lst.Clone())
    Call TestMax(lst.Clone())
    Call TestMin(lst.Clone())
    Call TestPop(lst.Clone())
    Call TestPush(lst.Clone())
    Call TestRemove(lst.Clone())
    Call TestRemoveAll(lst.Clone())
    Call TestRemoveAt(lst.Clone())
    Call TestReplace(lst.Clone())
    Call TestReplaceLast(lst.Clone())
    Call TestReverse(lst.Clone())
    Call TestRotate(lst.Clone())
    Call TestSample
    Call TestSetAt(lst.Clone())
    Call TestSetDataType(lst.Clone())
    Call TestShift(lst.Clone())
    Call TestShuffle
    Call TestSlice(lst.Clone())
    Call TestSort(lst.Clone())
    Call TestSplice(lst.Clone())
    Call TestSwap(lst.Clone())
    Call TestToArray
    Call TestToString(lst.Clone())
    Call TestUnique(lst.Clone())
    Call TestUnshift(lst.Clone())
    Call TestValuesAt(lst.Clone())

End Sub

Private Sub TestAdd()
    Dim lst1    As List

    Debug.Print "--- TestAdd ---"

    Set lst1 = New List

    On Error Resume Next
    Call lst1.Add(1, 1)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    Call lst1.Add(0, 1)
    Call PrintResult(lst1.ToString() = "(1)", 2)

    Call lst1.Add(0, 2)
    Call PrintResult(lst1.ToString() = "(2, 1)", 3)

    Call lst1.Add(1, 3)
    Call PrintResult(lst1.ToString() = "(2, 3, 1)", 4)

    Call lst1.Add(3, 4)
    Call PrintResult(lst1.ToString() = "(2, 3, 1, 4)", 5)

End Sub

Private Sub TestClear(ByRef lst As List)

    Debug.Print "--- TestClear ---"

    Call lst.Clear
    Call PrintResult(LangUtils.ToString(lst.ToString) = "()", 1)

End Sub

Private Sub TestClone()
    Dim lst1    As List
    Dim lst2    As List

    Debug.Print "--- TestClone ---"

    Set lst1 = New List

    Set lst2 = lst1.Clone()
    Call PrintResult(lst2.ToString() = "()", 1)

    Call lst1.Concat(1, 2, 3)
    Set lst2 = lst1.Clone()
    Call PrintResult(lst2.ToString() = "(1, 2, 3)", 2)

    Call lst1.Push(0)
    Call PrintResult(lst2.ToString() = "(1, 2, 3)", 3)

End Sub

Private Sub TestConcat()
    Dim lst1    As List
    Dim lst2    As List
    Dim lst3    As List
    Dim lst4    As List
    Dim lst5    As List

    Debug.Print "--- TestConcat ---"

    Set lst1 = New List
    Call lst1.SetDataType("Integer")
    Call lst1.Push(11)
    Call lst1.Push(12)
    Call lst1.Push(13)

    Set lst2 = New List
    Call lst2.SetDataType("Integer")
    Call lst2.Push(21)
    Call lst2.Push(22)
    Call lst2.Push(23)

    Set lst3 = New List
    Call lst3.SetDataType("Integer")
    Call lst3.Push(31)
    Call lst3.Push(32)
    Call lst3.Push(33)

    ' 引数指定なし
    Call lst1.Concat
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(11, 12, 13)", 1)

    ' 単一要素追加
    Call lst1.Concat(14)
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(11, 12, 13, 14)", 2)

    ' 複数要素追加
    Call lst1.Concat(15, 16)
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(11, 12, 13, 14, 15, 16)", 3)

    ' 配列追加
    Call lst1.Concat(Array(17, 18))
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(11, 12, 13, 14, 15, 16, 17, 18)", 4)

    ' リスト追加
    Call lst1.Concat(lst2)
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(11, 12, 13, 14, 15, 16, 17, 18, 21, 22, 23)", 5)

    Call lst1.Clear

    ' 複数リスト追加
    Call lst1.Concat(lst2, lst3)
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(21, 22, 23, 31, 32, 33)", 6)

    ' 単一要素、配列、リスト追加
    Call lst1.Concat(1, Array(2, 3), lst2)
    Call PrintResult(LangUtils.ToString(lst1.ToString) = "(21, 22, 23, 31, 32, 33, 1, 2, 3, 21, 22, 23)", 7)

    ' 複数階層のリスト、配列の場合
    Set lst4 = New List
    Call lst4.Push(Array(1, 2, 3))
    Call lst4.Push(lst2)

    Set lst5 = New List
    Call lst5.Concat(lst4, Array(Array(0, 1), Array(2, 3)))
    Call PrintResult(LangUtils.ToString(lst5.ToString) = "((1, 2, 3), (21, 22, 23), (0, 1), (2, 3))", 8)

End Sub

Private Sub TestContains(ByRef lst As List)

    Debug.Print "--- TestContains ---"

    Call PrintResult(lst.Contains(1), 1)

    Call PrintResult(Not lst.Contains(4), 2)

End Sub

Private Sub TestContainsAll(ByRef lst As List)

    Debug.Print "--- TestContainsAll ---"

    Call PrintResult(lst.ContainsAll(1, 2, 3), 1)

    Call PrintResult(Not lst.ContainsAll(2, 3, 4), 2)

End Sub

Private Sub TestContainsAny(ByRef lst As List)

    Debug.Print "--- TestContainsAny ---"

    Call PrintResult(lst.ContainsAny(1, 2, 3), 1)

    Call PrintResult(lst.ContainsAny(3, 4, 5), 2)

End Sub

Private Sub TestCount(ByRef lst As List)

    Debug.Print "--- TestCount ---"

    Call PrintResult(lst.Count(0) = 0, 1)
    Call PrintResult(lst.Count(1) = 1, 2)
    Call PrintResult(lst.Count(3) = 3, 3)

End Sub

Private Sub TestFirst(ByRef lst As List)
    Dim subLst      As List

    Debug.Print "--- TestFirst ---"

    Set subLst = New List

    Call PrintResult(lst.First() = 1, 1)

    Call subLst.Concat("A", "B", "C")
    Call lst.Unshift(subLst)

    Call PrintResult(lst.First().ToString() = "(A, B, C)", 2)

End Sub

Private Sub TestGetAt(ByRef lst As List)
    Dim subLst      As List
    Dim val         As Variant

    Debug.Print "--- TestGetAt ---"

    Set subLst = New List

    Call subLst.Concat("A", "B", "C")
    Call lst.Push(subLst)

    val = lst.GetAt(0)
    Call PrintResult(val = 1, 1)

    Set val = lst.GetAt(6)
    Call PrintResult(val.ToString() = "(A, B, C)", 2)

End Sub

Private Sub TestGetDataType(ByRef lst As List)

    Debug.Print "--- TestGetDataType ---"

    Call PrintResult(lst.GetDataType() = "", 1)

    Call lst.SetDataType("Integer")
    Call PrintResult(lst.GetDataType() = "Integer", 2)

    Call lst.SetDataType("")
    Call PrintResult(lst.GetDataType() = "", 3)

End Sub

Private Sub TestIndexOf(ByRef lst As List)

    Debug.Print "--- TestIndexOf ---"

    Call PrintResult(lst.IndexOf(2) = 1, 1)
    Call PrintResult(lst.IndexOf(2, 2) = 3, 2)

End Sub

Private Sub TestIndicesOf(ByRef lst As List)
    Dim indices()   As Long

    Debug.Print "--- TestIndicesOf ---"

    indices = lst.IndicesOf(3)
    Call PrintResult(LangUtils.ToString(indices) = "(2, 4, 5)", 1)

    indices = lst.IndicesOf(3, 3)
    Call PrintResult(LangUtils.ToString(indices) = "(4, 5)", 2)

End Sub

Private Sub TestIsEmptyList(ByRef lst As List)

    Debug.Print "--- TestIsEmptyList ---"

    Call PrintResult(Not lst.IsEmptyList(), 1)

    Call lst.Clear
    Call PrintResult(lst.IsEmptyList(), 2)

End Sub

Private Sub TestIsEqual(ByRef lst As List)
    Dim lst2    As List
    Dim subLst1 As List
    Dim subLst2 As List

    Debug.Print "--- TestIsEqual ---"

    Set lst2 = lst.Clone()

    Call PrintResult(lst.IsEqual(lst2), 1)

    Call lst.Push(Array("A"))
    Call PrintResult(Not lst.IsEqual(lst2), 2)

    Call lst2.Push(Array("A"))
    Call PrintResult(lst.IsEqual(lst2), 3)

    Set subLst1 = New List
    Call subLst1.Push("A")
    Set subLst2 = subLst1.Clone()
    Call lst.Push(subLst1)
    Call lst2.Push(subLst2)

    ' 入れ子になったListの要素までは比較できない
    Call PrintResult(Not lst.IsEqual(lst2), 4)

End Sub

Private Sub TestJoin()
    Dim lst1    As List
    Dim lst2    As List

    Debug.Print "--- TestJoin ---"

    Set lst1 = New List
    Set lst2 = New List

    Call PrintResult(lst1.Join() = "", 1)

    Call lst1.Concat(1, 2, 3)
    Call lst2.Concat("A", "B", "C")
    Call lst1.Push(lst2)

    Call PrintResult(lst1.Join() = "123(A, B, C)", 2)

    Call PrintResult(lst1.Join("-") = "1-2-3-(A, B, C)", 3)

End Sub

Private Sub TestLast(ByRef lst As List)
    Dim subLst      As List

    Debug.Print "--- TestLast ---"

    Call PrintResult(lst.Last() = 3, 1)

    Set subLst = New List
    Call subLst.Concat("A", "B", "C")
    Call lst.Push(subLst)

    Call PrintResult(lst.Last().ToString() = "(A, B, C)", 2)

End Sub

Private Sub TestLastIndexOf(ByRef lst As List)

    Debug.Print "--- TestLastIndexOf ---"

    Call PrintResult(lst.LastIndexOf(2) = 3, 1)
    Call PrintResult(lst.LastIndexOf(2, 2) = 1, 2)

End Sub

Private Sub TestLastIndicesOf(ByRef lst As List)
    Dim indices()   As Long

    Debug.Print "--- TestLastIndicesOf ---"

    indices = lst.LastIndicesOf(3)
    Call PrintResult(LangUtils.ToString(indices) = "(5, 4, 2)", 1)

    indices = lst.LastIndicesOf(3, 4)
    Call PrintResult(LangUtils.ToString(indices) = "(4, 2)", 2)

End Sub

Private Sub TestLength(ByRef lst As List)

    Debug.Print "--- TestLength ---"

    Call PrintResult(lst.Length() = 6, 1)

    Call lst.Clear
    Call PrintResult(lst.Length() = 0, 2)

End Sub

Private Sub TestMax(ByRef lst As List)

    Debug.Print "--- TestMax ---"

    Call PrintResult(lst.Max() = 3, 1)

End Sub

Private Sub TestMin(ByRef lst As List)

    Debug.Print "--- TestMin ---"

    Call PrintResult(lst.Min() = 1, 1)

End Sub

Private Sub TestPop(ByRef lst As List)
    Dim subLst  As List
    Dim val     As Variant

    Debug.Print "--- TestPop ---"

    Set subLst = New List
    Call lst.Push(subLst)

    Set val = lst.Pop
    Call PrintResult(val.ToString() = "()", "1-1")
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "1-2")
    val = lst.Pop
    Call PrintResult(val = 3, "2-1")
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3)", "2-2")

End Sub

Private Sub TestPush(ByRef lst As List)

    Debug.Print "--- TestPush ---"

    Call lst.Push(1)
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3, 1)", 1)
    Call lst.Push(2)
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3, 1, 2)", 2)

End Sub

Private Sub TestRemove(ByRef lst As List)

    Debug.Print "--- TestRemove ---"

    Call lst.Remove(3)
    Call PrintResult(lst.ToString() = "(1, 2, 2)", 1)

    Call lst.Remove(2, 1)
    Call PrintResult(lst.ToString() = "(1, 2)", 2)

End Sub

Private Sub TestRemoveAll(ByRef lst As List)

    Debug.Print "--- TestRemoveAll ---"

    Call lst.RemoveAll(1, 2, 3)
    Call PrintResult(lst.ToString() = "()", 1)

End Sub

Private Sub TestRemoveAt(ByRef lst As List)
    Dim val     As Variant

    Debug.Print "--- TestRemoveAt ---"

    On Error Resume Next
    Call lst.RemoveAt(6)
    Call PrintResult(Err.Number <> 0, 1)
    On Error GoTo 0

    val = lst.RemoveAt(2)
    Call PrintResult(lst.ToString() = "(1, 2, 2, 3, 3)", 2)

End Sub

Private Sub TestReplace(ByRef lst As List)
    Dim val     As Long

    Debug.Print "--- TestReplace ---"

    val = lst.Replace(3, 0)
    Call PrintResult(val = 3, 1)
    Call PrintResult(lst.ToString() = "(1, 2, 0, 2, 0, 0)", 2)

    val = lst.Replace(2, -1, 1)
    Call PrintResult(val = 1, 3)
    Call PrintResult(lst.ToString() = "(1, -1, 0, 2, 0, 0)", 4)

End Sub

Private Sub TestReplaceLast(ByRef lst As List)
    Dim val     As Long

    Debug.Print "--- TestReplaceLast ---"

    val = lst.ReplaceLast(3, 0)
    Call PrintResult(val = 3, 1)
    Call PrintResult(lst.ToString() = "(1, 2, 0, 2, 0, 0)", 2)

    val = lst.ReplaceLast(2, -1, 1)
    Call PrintResult(val = 1, 3)
    Call PrintResult(lst.ToString() = "(1, 2, 0, -1, 0, 0)", 4)

End Sub

Private Sub TestReverse(ByRef lst As List)

    Debug.Print "--- TestReverse ---"

    Call lst.Reverse
    Call PrintResult(lst.ToString() = "(3, 3, 2, 3, 2, 1)", 1)

End Sub

Private Sub TestRotate(ByRef lst As List)

    Debug.Print "--- TestRotate ---"

    Call lst.Rotate(2)
    Call PrintResult(lst.ToString() = "(3, 3, 1, 2, 3, 2)", 1)

    Call lst.Rotate(-3)
    Call PrintResult(lst.ToString() = "(2, 3, 2, 3, 3, 1)", 2)

End Sub

Private Sub TestSample()
    Dim lst1    As List
    Dim lst2    As List

    Debug.Print "--- TestSample ---"

    Set lst1 = New List
    Call lst1.Concat(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Set lst2 = lst1.Sample()
    Debug.Print lst2.ToString()

    Set lst2 = lst1.Sample(5)
    Debug.Print lst2.ToString()

    Set lst2 = lst1.Sample(6, False)
    Debug.Print lst2.ToString()

    Set lst2 = lst1.Sample(10, True)
    Debug.Print lst2.ToString()

End Sub

Private Sub TestSetAt(ByRef lst As List)
    Dim subLst  As List
    Dim val     As Variant

    Debug.Print "--- TestSetAt ---"

    Set subLst = New List
    Call lst.Push(subLst)

    val = lst.SetAt(1, 0)
    Call PrintResult(val = 2, "1-1")
    Call PrintResult(lst.ToString() = "(1, 0, 3, 2, 3, 3, ())", "1-2")

    Set val = lst.SetAt(6, 0)
    Call PrintResult(val.ToString() = "()", "2-1")
    Call PrintResult(lst.ToString() = "(1, 0, 3, 2, 3, 3, 0)", "2-2")

End Sub

Private Sub TestSetDataType(ByRef lst As List)
    Dim addLst  As List
    Dim lst1    As List

    Debug.Print "--- TestSetDataType ---"

    ' すでに要素が追加されている場合は変更不可、の確認
    Call PrintResult(lst.SetDataType("String") = False, "1-1")
    Call PrintResult(lst.GetDataType() = "", "1-2")

    Call lst.SetDataType("Integer")

    On Error Resume Next
    Call lst.Add(0, "1")
    Call PrintResult(Err.Number = 13, "2-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "2-2")

    On Error Resume Next
    Call lst.Concat(1, Array("2"))
    Call PrintResult(Err.Number = 13, "3-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "3-2")

    On Error Resume Next
    Call lst.Push("1")
    Call PrintResult(Err.Number = 13, "4-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "4-2")

    On Error Resume Next
    Call lst.Replace(1, "1")
    Call PrintResult(Err.Number = 13, "5-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "5-2")

    On Error Resume Next
    Call lst.Replace(1, "1")
    Call PrintResult(Err.Number = 13, "5-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "5-2")

    On Error Resume Next
    Call lst.ReplaceLast(1, "1")
    Call PrintResult(Err.Number = 13, "6-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "6-2")

    On Error Resume Next
    Call lst.SetAt(1, "1")
    Call PrintResult(Err.Number = 13, "7-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "7-2")

    Set addLst = New List
    Call addLst.Concat(4, 5, "6")

    On Error Resume Next
    Call lst.Splice(0, 0, addLst)
    Call PrintResult(Err.Number = 13, "8-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "8-2")

    On Error Resume Next
    Call lst.Unshift("1")
    Call PrintResult(Err.Number = 13, "9-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", "9-2")

    ' クラスも指定できることを確認
    Call lst.Clear
    Call lst.SetDataType("List")

    Set lst1 = New List
    Call lst1.Concat(1, 2, 3)
    Call lst.Push(lst1)

    On Error Resume Next
    Call lst.Push(1)
    Call PrintResult(Err.Number = 13, "10-1")
    On Error GoTo 0
    Call PrintResult(lst.ToString() = "((1, 2, 3))", "10-2")

End Sub

Private Sub TestShift(ByRef lst As List)
    Dim subLst  As List
    Dim val     As Variant

    Debug.Print "--- TestShift ---"

    Set subLst = New List
    Call lst.Add(1, subLst)

    val = lst.Shift
    Call PrintResult(val = 1, "1-1")
    Call PrintResult(lst.ToString() = "((), 2, 3, 2, 3, 3)", "1-2")

    Set val = lst.Shift
    Call PrintResult(val.ToString() = "()", "2-1")
    Call PrintResult(lst.ToString() = "(2, 3, 2, 3, 3)", "2-2")

End Sub

Private Sub TestShuffle()
    Dim lst     As List

    Debug.Print "--- TestShuffle ---"

    Set lst = New List

    Call lst.Shuffle
    Call PrintResult(lst.ToString() = "()", 1)

    Call lst.Concat(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Call lst.Shuffle
    Debug.Print lst.ToString()

End Sub

Private Sub TestSlice(ByRef lst As List)
    Dim result  As List

    Debug.Print "--- TestSlice ---"

    Set result = lst.Slice(1, 4)
    Call PrintResult(result.ToString() = "(2, 3, 2)", 1)

    Set result = lst.Slice(4)
    Call PrintResult(result.ToString() = "(3, 3)", 2)

End Sub

Private Sub TestSort(ByRef lst As List)

    Debug.Print "--- TestSort ---"

    Call lst.Sort
    Call PrintResult(lst.ToString() = "(1, 2, 2, 3, 3, 3)", 1)

End Sub

Private Sub TestSplice(ByRef lst As List)
    Dim addLst  As List

    Debug.Print "--- TestSplice ---"

    Set addLst = New List
    Call addLst.Concat(4, 5)

    Call lst.Splice(2, 2)
    Call PrintResult(lst.ToString() = "(1, 2, 3, 3)", 1)

    Call lst.Splice(0, 0, addLst)
    Call PrintResult(lst.ToString() = "(4, 5, 1, 2, 3, 3)", 2)

End Sub

Private Sub TestSwap(ByRef lst As List)
    Dim arr()   As Variant

    Debug.Print "--- TestSwap ---"

    Call lst.Swap(1, 2)
    Call PrintResult(lst.ToString() = "(1, 3, 2, 2, 3, 3)", 1)

End Sub

Private Sub TestToArray()
    Dim lst     As List
    Dim arr()   As Variant

    Debug.Print "--- TestToArray ---"

    Set lst = New List

    arr = lst.ToArray()
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)

    Call lst.Concat(1, 2, 3)

    arr = lst.ToArray()
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 2)

    arr(0) = 0
    Call PrintResult(lst.ToString() = "(1, 2, 3)", 3)

End Sub

Private Sub TestToString(ByRef lst As List)
    Dim subLst  As List

    Debug.Print "--- TestToString ---"

    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3)", 1)

    Set subLst = New List
    Call subLst.Concat("A", "B", "C")
    Call lst.Push(subLst)

    Call PrintResult(lst.ToString() = "(1, 2, 3, 2, 3, 3, (A, B, C))", 2)

End Sub

Private Sub TestUnique(ByRef lst As List)

    Debug.Print "--- TestUnique ---"

    Call lst.Unique
    Call PrintResult(lst.ToString() = "(1, 2, 3)", 1)

End Sub

Private Sub TestUnshift(ByRef lst As List)

    Debug.Print "--- TestUnshift ---"

    Call lst.Unshift(1)
    Call PrintResult(lst.ToString() = "(1, 1, 2, 3, 2, 3, 3)", 1)
    Call lst.Unshift(2)
    Call PrintResult(lst.ToString() = "(2, 1, 1, 2, 3, 2, 3, 3)", 2)

End Sub

Private Sub TestValuesAt(ByRef lst As List)
    Dim result  As List

    Debug.Print "--- TestValuesAt ---"

    Set result = lst.ValuesAt(0, 2, 5)
    Call PrintResult(result.ToString() = "(1, 3, 3)", 1)

End Sub

