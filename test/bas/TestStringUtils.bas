Attribute VB_Name = "TestStringUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : StringUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)

    Debug.Print "=== TestStringUtils ==="

    Call TestAppendIfMissing
    Call TestContains
    Call TestContainsAll
    Call TestContainsAny
    Call TestCount
    Call TestEndsWith
    Call TestEndsWithAny
    Call TestEquals
    Call TestEqualsAny
    Call TestFirstNotBlank
    Call TestIndexOf
    Call TestIndexOfAny
    Call TestIsAllBlank
    Call TestIsAllEmpty
    Call TestIsAlpha
    Call TestIsAlphaDigit
    Call TestIsAnyBlank
    Call TestIsAnyEmpty
    Call TestIsBlank
    Call TestIsDigit
    Call TestLastIndexOf
    Call TestLastIndexOfAny
    Call TestLeftBefore
    Call TestLTrim
    Call TestMidBetween
    Call TestOverlay
    Call TestPartition
    Call TestPrependIfMissing
    Call TestRemove
    Call TestRemoveEnd
    Call TestRemoveStart
    Call TestReplace
    Call TestReverse
    Call TestRightAfter
    Call TestRotate
    Call TestRTrim
    Call TestSplitByAlpha
    Call TestSplitByBlank
    Call TestSplitByChars
    Call TestSplitByDigit
    Call TestSplitByNewline
    Call TestStartsWith
    Call TestStartsWithAny
    Call TestTrim

End Sub

Private Sub TestAppendIfMissing()
    Dim str     As String

    Debug.Print "--- TestAppendIfMissing ---"

    Call PrintResult(StringUtils.AppendIfMissing(str, "") = "", 1)
    Call PrintResult(StringUtils.AppendIfMissing(str, "abc") = "abc", 2)

    str = "abcdef"
    Call PrintResult(StringUtils.AppendIfMissing(str, "g") = "abcdefg", 3)
    Call PrintResult(StringUtils.AppendIfMissing(str, "abcde") = "abcdefabcde", 4)
    Call PrintResult(StringUtils.AppendIfMissing(str, "DEF") = "abcdefDEF", 5)
    Call PrintResult(StringUtils.AppendIfMissing(str, "") = "abcdef", 6)
    Call PrintResult(StringUtils.AppendIfMissing(str, "f") = "abcdef", 7)
    Call PrintResult(StringUtils.AppendIfMissing(str, "abcdef") = "abcdef", 8)
    Call PrintResult(StringUtils.AppendIfMissing(str, "DEF", True) = "abcdef", 9)

End Sub

Private Sub TestContains()
    Dim str     As String

    Debug.Print "--- TestContains ---"

    Call PrintResult(Not StringUtils.Contains(str, ""), 1)
    Call PrintResult(Not StringUtils.Contains(str, "a"), 2)

    str = "abcde"
    Call PrintResult(StringUtils.Contains(str, "a"), 3)
    Call PrintResult(StringUtils.Contains(str, "c"), 4)
    Call PrintResult(StringUtils.Contains(str, "e"), 5)
    Call PrintResult(StringUtils.Contains(str, "abcde"), 6)
    Call PrintResult(Not StringUtils.Contains(str, "f"), 7)
    Call PrintResult(Not StringUtils.Contains(str, "abde"), 8)
    Call PrintResult(Not StringUtils.Contains(str, "A"), 9)
    Call PrintResult(StringUtils.Contains(str, "A", True), 10)

End Sub

Private Sub TestContainsAll()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestContainsAll ---"

    Call PrintResult(Not StringUtils.ContainsAll(str, params), 1)

    params = ArrayUtils.CStrArray("", "a")
    Call PrintResult(Not StringUtils.ContainsAll(str, "", "a"), 2)
    Call PrintResult(Not StringUtils.ContainsAll(str, params), 3)

    str = "abcde"
    params = ArrayUtils.CStrArray("a", "abc", "abcde")
    Call PrintResult(StringUtils.ContainsAll(str, "a"), 4)
    Call PrintResult(StringUtils.ContainsAll(str, "a", "abc", "abcde"), 5)
    Call PrintResult(StringUtils.ContainsAll(str, params), 6)

    Call PrintResult(Not StringUtils.ContainsAll(str, "f"), 7)
    Call PrintResult(Not StringUtils.ContainsAll(str, "", "a"), 8)
    Call PrintResult(Not StringUtils.ContainsAll(str, "A", "a", "abc"), 9)

End Sub

Private Sub TestContainsAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestContainsAny ---"

    Call PrintResult(Not StringUtils.ContainsAny(str, params), 1)

    params = ArrayUtils.CStrArray("", "a")
    Call PrintResult(Not StringUtils.ContainsAny(str, "", "a"), 2)
    Call PrintResult(Not StringUtils.ContainsAny(str, params), 3)

    str = "abcde"
    params = ArrayUtils.CStrArray("a", "abc", "abcde")
    Call PrintResult(StringUtils.ContainsAny(str, "a"), 4)
    Call PrintResult(StringUtils.ContainsAny(str, "a", "abc", "abcde"), 5)
    Call PrintResult(StringUtils.ContainsAny(str, params), 6)

    Call PrintResult(Not StringUtils.ContainsAny(str, "f"), 7)
    Call PrintResult(StringUtils.ContainsAny(str, "", "a"), 8)
    Call PrintResult(StringUtils.ContainsAny(str, "A", "f", "abc"), 9)
    Call PrintResult(Not StringUtils.ContainsAny(str, "A", "f", "ABC"), 10)

End Sub

Private Sub TestCount()
    Dim str         As String

    Debug.Print "--- TestCount ---"

    Call PrintResult(StringUtils.Count(str, "") = 0, 1)
    Call PrintResult(StringUtils.Count(str, "a") = 0, 2)

    str = "abaabaabaa"
    Call PrintResult(StringUtils.Count(str, "A") = 0, 3)
    Call PrintResult(StringUtils.Count(str, "a") = 7, 4)
    Call PrintResult(StringUtils.Count(str, "abaabaabaa") = 1, 5)
    Call PrintResult(StringUtils.Count(str, "aba") = 3, 6)
    Call PrintResult(StringUtils.Count(str, "abaa") = 2, 7)

End Sub

Private Sub TestEndsWith()
    Dim str     As String

    Debug.Print "--- TestEndsWith ---"

    str = ""
    Call PrintResult(Not StringUtils.EndsWith(str, ""), 1)
    Call PrintResult(Not StringUtils.EndsWith(str, "a"), 2)

    str = "abcdef"
    Call PrintResult(StringUtils.EndsWith(str, "f"), 3)
    Call PrintResult(StringUtils.EndsWith(str, "abcdef"), 4)
    Call PrintResult(Not StringUtils.EndsWith(str, "e"), 5)
    Call PrintResult(Not StringUtils.EndsWith(str, "DEF"), 6)
    Call PrintResult(StringUtils.EndsWith(str, "DEF", True), 7)

End Sub

Private Sub TestEndsWithAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestEndsWithAny ---"

    str = ""
    Call PrintResult(Not StringUtils.EndsWithAny(str, ""), 1)
    Call PrintResult(Not StringUtils.EndsWithAny(str, params), 2)

    str = "abcdef"
    params = ArrayUtils.CStrArray("f", "ef", "def")
    Call PrintResult(Not StringUtils.EndsWithAny(str, ""), 3)
    Call PrintResult(StringUtils.EndsWithAny(str, "def"), 4)
    Call PrintResult(StringUtils.EndsWithAny(str, params), 5)
    Call PrintResult(StringUtils.EndsWithAny(str, "abc", "def", "g"), 6)
    Call PrintResult(Not StringUtils.EndsWithAny(str, "", "g", "1"), 7)

End Sub

Private Sub TestEquals()

    Debug.Print "--- TestEquals ---"

    Call PrintResult(StringUtils.Equals("", ""), 1)
    Call PrintResult(StringUtils.Equals("abc", "abc"), 2)
    Call PrintResult(Not StringUtils.Equals("abc", "ab"), 3)
    Call PrintResult(Not StringUtils.Equals("abc", "ABC"), 4)
    Call PrintResult(StringUtils.Equals("abc", "ABC", True), 5)
    Call PrintResult(Not StringUtils.Equals("あ", "ア", True), 6)

End Sub

Private Sub TestEqualsAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestEqualsAny ---"

    Call PrintResult(StringUtils.EqualsAny(str, ""), 1)
    Call PrintResult(Not StringUtils.EqualsAny(str, params), 2)

    str = "abc"
    params = ArrayUtils.CStrArray("abc", "def", "123")
    Call PrintResult(StringUtils.EqualsAny(str, "abc"), 3)
    Call PrintResult(Not StringUtils.EqualsAny(str, "ABC"), 4)
    Call PrintResult(StringUtils.EqualsAny(str, "", "abcd", "abc"), 5)
    Call PrintResult(StringUtils.EqualsAny(str, params), 6)
    Call PrintResult(Not StringUtils.EqualsAny(str, "abcd", "def", "123"), 7)

End Sub

Private Sub TestFirstNotBlank()
    Dim params()    As String

    Debug.Print "--- TestFirstNotBlank ---"

    Call PrintResult(StringUtils.FirstNotBlank(params) = "", 1)

    params = ArrayUtils.CStrArray("", " 　", "abc")
    Call PrintResult(StringUtils.FirstNotBlank("") = "", 2)
    Call PrintResult(StringUtils.FirstNotBlank("abc", "123", " 　") = "abc", 3)
    Call PrintResult(StringUtils.FirstNotBlank(" 　", "123", "abc") = "123", 4)
    Call PrintResult(StringUtils.FirstNotBlank(" 　", vbCrLf, vbTab) = "", 5)
    Call PrintResult(StringUtils.FirstNotBlank(params) = "abc", 6)

End Sub

Private Sub TestIndexOf()
    Dim str     As String

    Debug.Print "--- TestIndexOf ---"

    str = ""
    Call PrintResult(StringUtils.IndexOf(str, "") = 0, 1)
    Call PrintResult(StringUtils.IndexOf(str, "a") = 0, 2)

    str = "abcabcABC"
    Call PrintResult(StringUtils.IndexOf(str, "") = 0, 3)
    Call PrintResult(StringUtils.IndexOf(str, "abc") = 1, 4)
    Call PrintResult(StringUtils.IndexOf(str, "ABC") = 7, 5)
    Call PrintResult(StringUtils.IndexOf(str, "C") = 9, 6)
    Call PrintResult(StringUtils.IndexOf(str, "abcabcABC") = 1, 7)
    Call PrintResult(StringUtils.IndexOf(str, "f") = 0, 8)
    Call PrintResult(StringUtils.IndexOf(str, "a", 2) = 4, 9)
    Call PrintResult(StringUtils.IndexOf(str, "ABC", 3, True) = 4, 10)

End Sub

Private Sub TestIndexOfAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestIndexOfAny ---"

    str = ""
    Call PrintResult(StringUtils.IndexOfAny(str, "") = 0, 1)
    Call PrintResult(StringUtils.IndexOfAny(str, params) = 0, 2)

    str = "abcabcABC"
    params = ArrayUtils.CStrArray("", "c", "d")
    Call PrintResult(StringUtils.IndexOfAny(str, "") = 0, 3)
    Call PrintResult(StringUtils.IndexOfAny(str, "abc", "b", "ABC") = 1, 4)
    Call PrintResult(StringUtils.IndexOfAny(str, params) = 3, 5)
    Call PrintResult(StringUtils.IndexOfAny(str, "C", "d") = 9, 6)
    Call PrintResult(StringUtils.IndexOfAny(str, "d", "1", "abcd") = 0, 7)

End Sub

Private Sub TestIsAllBlank()
    Dim params()    As String

    Debug.Print "--- TestIsAllBlank ---"

    Call PrintResult(Not StringUtils.IsAllBlank(params), 1)

    params = ArrayUtils.CStrArray("", " 　", vbTab, vbCrLf, vbLf)
    Call PrintResult(StringUtils.IsAllBlank(" 　"), 2)
    Call PrintResult(StringUtils.IsAllBlank("", vbTab), 3)
    Call PrintResult(StringUtils.IsAllBlank(params), 4)

    Call PrintResult(Not StringUtils.IsAllBlank("a"), 5)
    Call PrintResult(Not StringUtils.IsAllBlank(vbCrLf, "a", vbLf), 6)

End Sub

Private Sub TestIsAllEmpty()
    Dim params()    As String

    Debug.Print "--- TestIsAllEmpty ---"

    Call PrintResult(Not StringUtils.IsAllEmpty(params), 1)

    params = ArrayUtils.CStrArray("", "", "")
    Call PrintResult(StringUtils.IsAllEmpty(""), 2)
    Call PrintResult(StringUtils.IsAllEmpty("", ""), 3)
    Call PrintResult(StringUtils.IsAllEmpty(params), 4)

    Call PrintResult(Not StringUtils.IsAllEmpty("a"), 5)
    Call PrintResult(Not StringUtils.IsAllEmpty("", "a", ""), 6)

End Sub

Private Sub TestIsAlpha()
    Dim i       As Long

    Debug.Print "--- TestIsAlpha ---"

    Call PrintResult(Not StringUtils.IsAlpha(""), 1)
    Call PrintResult(Not StringUtils.IsAlpha(Chr(64)), 2)
    For i = 65 To 90
        Call PrintResult(StringUtils.IsAlpha(Chr(i)), 3)
    Next
    Call PrintResult(Not StringUtils.IsAlpha(Chr(91)), 4)
    Call PrintResult(Not StringUtils.IsAlpha(Chr(96)), 5)
    For i = 97 To 122
        Call PrintResult(StringUtils.IsAlpha(Chr(i)), 6)
    Next
    Call PrintResult(Not StringUtils.IsAlpha(Chr(123)), 7)
    Call PrintResult(StringUtils.IsAlpha("AzByCxDwEvFuGtHsIrJqKpLoMnNmOlPkQjRiShTgUfVeWdXcYbZa"), 8)
    Call PrintResult(Not StringUtils.IsAlpha("1"), 9)
    Call PrintResult(Not StringUtils.IsAlpha("あ"), 10)

End Sub

Private Sub TestIsAlphaDigit()
    Dim i       As Long

    Debug.Print "--- TestIsAlphaDigit ---"

    Call PrintResult(Not StringUtils.IsAlphaDigit(""), 1)
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(47)), 2)
    For i = 48 To 57
        Call PrintResult(StringUtils.IsAlphaDigit(Chr(i)), 3)
    Next
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(58)), 4)
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(64)), 5)
    For i = 65 To 90
        Call PrintResult(StringUtils.IsAlphaDigit(Chr(i)), 6)
    Next
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(91)), 7)
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(96)), 8)
    For i = 97 To 122
        Call PrintResult(StringUtils.IsAlphaDigit(Chr(i)), 9)
    Next
    Call PrintResult(Not StringUtils.IsAlphaDigit(Chr(123)), 10)
    Call PrintResult(StringUtils.IsAlphaDigit("01234AzByCxDwEvFuGtHsIrJqKpLoMnNmOlPkQjRiShTgUfVeWdXcYbZa56789"), 11)
    Call PrintResult(Not StringUtils.IsAlphaDigit("-"), 12)
    Call PrintResult(Not StringUtils.IsAlphaDigit("あ"), 13)

End Sub

Private Sub TestIsAnyBlank()
    Dim params()    As String

    Debug.Print "--- TestIsAnyBlank ---"

    Call PrintResult(Not StringUtils.IsAnyBlank(params), 1)

    params = ArrayUtils.CStrArray(vbTab, "a", vbCrLf)
    Call PrintResult(StringUtils.IsAnyBlank(""), 2)
    Call PrintResult(StringUtils.IsAnyBlank(" 　", "a"), 3)
    Call PrintResult(StringUtils.IsAnyBlank(params), 4)

    Call PrintResult(Not StringUtils.IsAnyBlank("a"), 5)
    Call PrintResult(Not StringUtils.IsAnyBlank("a", "b", "c"), 6)

End Sub

Private Sub TestIsAnyEmpty()
    Dim params()    As String

    Debug.Print "--- TestIsAnyEmpty ---"

    Call PrintResult(Not StringUtils.IsAnyEmpty(params), 1)

    params = ArrayUtils.CStrArray("", "a", "")
    Call PrintResult(StringUtils.IsAnyEmpty(""), 2)
    Call PrintResult(StringUtils.IsAnyEmpty("", "a"), 3)
    Call PrintResult(StringUtils.IsAnyEmpty(params), 4)

    Call PrintResult(Not StringUtils.IsAnyEmpty("a"), 5)
    Call PrintResult(Not StringUtils.IsAnyEmpty("a", "b", "c"), 6)

End Sub

Private Sub TestIsBlank()
    Dim str     As String

    Debug.Print "--- TestIsBlank ---"

    Call PrintResult(StringUtils.IsBlank(str), 1)

    str = ""
    Call PrintResult(StringUtils.IsBlank(str), 2)

    str = " "       ' 半角スペース
    Call PrintResult(StringUtils.IsBlank(str), 3)

    str = "　"      ' 全角スペース
    Call PrintResult(StringUtils.IsBlank(str), 4)

    str = vbTab     ' タブ
    Call PrintResult(StringUtils.IsBlank(str), 5)

    str = vbCrLf    ' 改行文字
    Call PrintResult(StringUtils.IsBlank(str), 6)

    str = vbLf      ' 改行文字
    Call PrintResult(StringUtils.IsBlank(str), 7)

    str = " 　" & vbTab & vbCrLf & vbLf     ' 全空白文字混合
    Call PrintResult(StringUtils.IsBlank(str), 8)

    str = " a　"
    Call PrintResult(Not StringUtils.IsBlank(str), 9)

End Sub

Private Sub TestIsDigit()
    Dim i       As Long

    Debug.Print "--- TestIsDigit ---"

    Call PrintResult(Not StringUtils.IsDigit(""), 1)
    For i = 0 To 9
        Call PrintResult(StringUtils.IsDigit(CStr(i)), 2)
    Next
    Call PrintResult(StringUtils.IsDigit("9876543210"), 3)
    Call PrintResult(Not StringUtils.IsDigit("-1"), 4)
    Call PrintResult(Not StringUtils.IsDigit("２"), 5)
    Call PrintResult(Not StringUtils.IsDigit("a"), 6)

End Sub

Private Sub TestLastIndexOf()
    Dim str     As String

    Debug.Print "--- TestLastIndexOf ---"

    str = ""
    Call PrintResult(StringUtils.LastIndexOf(str, "") = 0, 1)

    str = "abcabcABC"
    Call PrintResult(StringUtils.LastIndexOf(str, "") = 0, 2)
    Call PrintResult(StringUtils.LastIndexOf(str, "abc") = 4, 3)
    Call PrintResult(StringUtils.LastIndexOf(str, "ABC") = 7, 4)
    Call PrintResult(StringUtils.LastIndexOf(str, "C") = 9, 5)
    Call PrintResult(StringUtils.LastIndexOf(str, "abcabcABC") = 1, 6)
    Call PrintResult(StringUtils.LastIndexOf(str, "f") = 0, 7)
    Call PrintResult(StringUtils.LastIndexOf(str, "a", 2) = 1, 8)
    Call PrintResult(StringUtils.LastIndexOf(str, "abc", 2) = 0, 9)
    Call PrintResult(StringUtils.LastIndexOf(str, "abc", -1, True) = 7, 10)

End Sub

Private Sub TestLastIndexOfAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestLastIndexOfAny ---"

    str = ""
    Call PrintResult(StringUtils.LastIndexOfAny(str, "") = 0, 1)
    Call PrintResult(StringUtils.LastIndexOfAny(str, params) = 0, 2)

    str = "abcabcABC"
    params = ArrayUtils.CStrArray("", "abc", "b")
    Call PrintResult(StringUtils.LastIndexOfAny(str, "") = 0, 3)
    Call PrintResult(StringUtils.LastIndexOfAny(str, "abc", "b", "ABC") = 7, 4)
    Call PrintResult(StringUtils.LastIndexOfAny(str, params) = 5, 5)
    Call PrintResult(StringUtils.LastIndexOfAny(str, "C", "d") = 9, 6)
    Call PrintResult(StringUtils.LastIndexOfAny(str, "d", "1", "abcd") = 0, 7)

End Sub

Private Sub TestLeftBefore()
    Dim str     As String

    Debug.Print "--- TestLeftBefore ---"

    Call PrintResult(StringUtils.LeftBefore(str, "") = "", 1)
    Call PrintResult(StringUtils.LeftBefore(str, "a") = "", 2)

    str = "abcabcd"
    Call PrintResult(StringUtils.LeftBefore(str, "a") = "", 3)
    Call PrintResult(StringUtils.LeftBefore(str, "b") = "a", 4)
    Call PrintResult(StringUtils.LeftBefore(str, "d") = "abcabc", 5)
    Call PrintResult(StringUtils.LeftBefore(str, "CA") = "", 6)
    Call PrintResult(StringUtils.LeftBefore(str, "CA", True) = "ab", 7)

End Sub

Private Sub TestLTrim()
    Dim str     As String

    Debug.Print "--- TestLTrim ---"

    str = ""
    Call PrintResult(StringUtils.LTrim(str) = "", 1)

    str = "ab c"
    Call PrintResult(StringUtils.LTrim(str) = "ab c", 2)

    str = " ab c "
    Call PrintResult(StringUtils.LTrim(str) = "ab c ", 3)

    str = " 　" & vbTab & vbCrLf & vbLf & "ab c "
    Call PrintResult(StringUtils.LTrim(str) = "ab c ", 4)

    str = " a"
    Call PrintResult(StringUtils.LTrim(str) = "a", 5)

    str = " 　"
    Call PrintResult(StringUtils.LTrim(str) = "", 6)

End Sub

Private Sub TestMidBetween()
    Dim str     As String

    Debug.Print "--- TestMidBetween ---"

    Call PrintResult(StringUtils.MidBetween(str, "", "") = "", 1)
    Call PrintResult(StringUtils.MidBetween(str, "a", "a") = "", 2)

    str = "abcabcdef"
    Call PrintResult(StringUtils.MidBetween(str, "a", "b") = "", 3)
    Call PrintResult(StringUtils.MidBetween(str, "a", "a") = "bc", 4)
    Call PrintResult(StringUtils.MidBetween(str, "b", "b") = "ca", 5)
    Call PrintResult(StringUtils.MidBetween(str, "c", "f") = "abcde", 6)
    Call PrintResult(StringUtils.MidBetween(str, "a", "f") = "bcabcde", 7)
    Call PrintResult(StringUtils.MidBetween(str, "f", "a") = "", 8)
    Call PrintResult(StringUtils.MidBetween(str, "A", "C") = "", 9)
    Call PrintResult(StringUtils.MidBetween(str, "A", "C", True) = "b", 10)

End Sub

Private Sub TestOverlay()
    Dim str     As String

    Debug.Print "--- TestOverlay ---"

    Call PrintResult(StringUtils.Overlay(str, "", 1, 2) = "", 1)
    Call PrintResult(StringUtils.Overlay(str, "a", 1, 2) = "", 2)

    str = "abcdef"
    Call PrintResult(StringUtils.Overlay(str, "", 3, 4) = "abef", 3)
    Call PrintResult(StringUtils.Overlay(str, "123", 1, 1) = "123bcdef", 4)
    Call PrintResult(StringUtils.Overlay(str, "123", 6, 6) = "abcde123", 5)
    Call PrintResult(StringUtils.Overlay(str, "123", 2, 5) = "a123f", 6)
    Call PrintResult(StringUtils.Overlay(str, "123", 1, 6) = "123", 7)
    Call PrintResult(StringUtils.Overlay(str, "123", 0, 2) = "123cdef", 8)
    Call PrintResult(StringUtils.Overlay(str, "123", 0, 0) = "abcdef", 9)
    Call PrintResult(StringUtils.Overlay(str, "123", 4, 7) = "abc123", 10)
    Call PrintResult(StringUtils.Overlay(str, "123", 7, 7) = "abcdef", 11)
    Call PrintResult(StringUtils.Overlay(str, "123", 5, 2) = "abcdef", 12)

End Sub

Private Sub TestPartition()
    Dim arr()   As String

    Debug.Print "--- TestPartition ---"

    arr = StringUtils.Partition("", "")
    Call PrintResult(LangUtils.ToString(arr) = "(, , )", 1)

    arr = StringUtils.Partition("abc", "")
    Call PrintResult(LangUtils.ToString(arr) = "(, , abc)", 2)

    arr = StringUtils.Partition("abcdefg", "abc")
    Call PrintResult(LangUtils.ToString(arr) = "(, abc, defg)", 3)

    arr = StringUtils.Partition("abcdefg", "cde")
    Call PrintResult(LangUtils.ToString(arr) = "(ab, cde, fg)", 4)

    arr = StringUtils.Partition("abcdefg", "efg")
    Call PrintResult(LangUtils.ToString(arr) = "(abcd, efg, )", 5)

    arr = StringUtils.Partition("abcdefg", "ABC")
    Call PrintResult(LangUtils.ToString(arr) = "(abcdefg, , )", 6)

    arr = StringUtils.Partition("abcdefg", "ABC", 1, True)
    Call PrintResult(LangUtils.ToString(arr) = "(, abc, defg)", 7)

End Sub

Private Sub TestPrependIfMissing()
    Dim str     As String

    Debug.Print "--- TestPrependIfMissing ---"

    Call PrintResult(StringUtils.PrependIfMissing(str, "") = "", 1)
    Call PrintResult(StringUtils.PrependIfMissing(str, "abc") = "abc", 2)

    str = "abcdef"
    Call PrintResult(StringUtils.PrependIfMissing(str, "1") = "1abcdef", 3)
    Call PrintResult(StringUtils.PrependIfMissing(str, "bcdef") = "bcdefabcdef", 4)
    Call PrintResult(StringUtils.PrependIfMissing(str, "ABC") = "ABCabcdef", 5)
    Call PrintResult(StringUtils.PrependIfMissing(str, "") = "abcdef", 6)
    Call PrintResult(StringUtils.PrependIfMissing(str, "a") = "abcdef", 7)
    Call PrintResult(StringUtils.PrependIfMissing(str, "abcdef") = "abcdef", 8)
    Call PrintResult(StringUtils.PrependIfMissing(str, "ABC", True) = "abcdef", 9)

End Sub

Private Sub TestRemove()
    Dim str     As String

    Debug.Print "--- TestRemove ---"

    Call PrintResult(StringUtils.Remove(str, "") = "", 1)
    Call PrintResult(StringUtils.Remove(str, "a") = "", 2)

    str = "abc1abc2abc3"
    Call PrintResult(StringUtils.Remove(str, "abc") = "123", 3)
    Call PrintResult(StringUtils.Remove(str, "abcd") = "abc1abc2abc3", 4)
    Call PrintResult(StringUtils.Remove(str, "ABC") = "abc1abc2abc3", 5)
    Call PrintResult(StringUtils.Remove(str, "abc", 2) = "abc123", 6)
    Call PrintResult(StringUtils.Remove(str, "abc", 13) = "abc1abc2abc3", 7)
    Call PrintResult(StringUtils.Remove(str, "abc", 1, 0) = "123", 8)
    Call PrintResult(StringUtils.Remove(str, "abc", 1, 1) = "1abc2abc3", 9)
    Call PrintResult(StringUtils.Remove(str, "abc", 2, 3) = "abc123", 10)
    Call PrintResult(StringUtils.Remove(str, "ABC", 1, 0, True) = "123", 11)

End Sub

Private Sub TestRemoveEnd()
    Dim str     As String

    Debug.Print "--- TestRemoveEnd ---"

    Call PrintResult(StringUtils.RemoveEnd(str, "") = "", 1)
    Call PrintResult(StringUtils.RemoveEnd(str, "abc") = "", 2)

    str = "abcdef"
    Call PrintResult(StringUtils.RemoveEnd(str, "f") = "abcde", 3)
    Call PrintResult(StringUtils.RemoveEnd(str, "abcdef") = "", 4)
    Call PrintResult(StringUtils.RemoveEnd(str, "abcdefg") = "abcdef", 5)
    Call PrintResult(StringUtils.RemoveEnd(str, "DEF") = "abcdef", 6)
    Call PrintResult(StringUtils.RemoveEnd(str, "DEF", True) = "abc", 7)

End Sub

Private Sub TestRemoveStart()
    Dim str     As String

    Debug.Print "--- TestRemoveStart ---"

    Call PrintResult(StringUtils.RemoveStart(str, "") = "", 1)
    Call PrintResult(StringUtils.RemoveStart(str, "abc") = "", 2)

    str = "abcdef"
    Call PrintResult(StringUtils.RemoveStart(str, "a") = "bcdef", 3)
    Call PrintResult(StringUtils.RemoveStart(str, "abcdef") = "", 4)
    Call PrintResult(StringUtils.RemoveStart(str, "abcdefg") = "abcdef", 5)
    Call PrintResult(StringUtils.RemoveStart(str, "ABC") = "abcdef", 6)
    Call PrintResult(StringUtils.RemoveStart(str, "ABC", True) = "def", 7)

End Sub

Private Sub TestReplace()
    Dim str     As String

    Debug.Print "--- TestReplace ---"

    Call PrintResult(StringUtils.Replace(str, "", "abc") = "", 1)
    Call PrintResult(StringUtils.Replace(str, "a", "abc") = "", 2)

    str = "abc1abc2abc3"
    Call PrintResult(StringUtils.Replace(str, "abc", "def") = "def1def2def3", 3)
    Call PrintResult(StringUtils.Replace(str, "abcd", "def") = "abc1abc2abc3", 4)
    Call PrintResult(StringUtils.Replace(str, "ABC", "def") = "abc1abc2abc3", 5)
    Call PrintResult(StringUtils.Replace(str, "abc", "def", 2) = "abc1def2def3", 6)
    Call PrintResult(StringUtils.Replace(str, "abc", "def", 13) = "abc1abc2abc3", 7)
    Call PrintResult(StringUtils.Replace(str, "abc", "def", 1, 0) = "def1def2def3", 8)
    Call PrintResult(StringUtils.Replace(str, "abc", "def", 1, 1) = "def1abc2abc3", 9)
    Call PrintResult(StringUtils.Replace(str, "abc", "def", 2, 3) = "abc1def2def3", 10)
    Call PrintResult(StringUtils.Replace(str, "ABC", "def", 1, 0, True) = "def1def2def3", 11)

End Sub

Private Sub TestReverse()

    Debug.Print "--- TestReverse ---"

    Call PrintResult(StringUtils.Reverse("") = "", 1)
    Call PrintResult(StringUtils.Reverse("a") = "a", 2)
    Call PrintResult(StringUtils.Reverse("abcdef") = "fedcba", 3)
    Call PrintResult(StringUtils.Reverse("12345") = "54321", 4)

End Sub

Private Sub TestRightAfter()
    Dim str     As String

    Debug.Print "--- TestRightAfter ---"

    Call PrintResult(StringUtils.RightAfter(str, "") = "", 1)
    Call PrintResult(StringUtils.RightAfter(str, "a") = "", 2)

    str = "abcabcd"
    Call PrintResult(StringUtils.RightAfter(str, "d") = "", 3)
    Call PrintResult(StringUtils.RightAfter(str, "cabc") = "d", 4)
    Call PrintResult(StringUtils.RightAfter(str, "a") = "bcabcd", 5)
    Call PrintResult(StringUtils.RightAfter(str, "CA") = "", 6)
    Call PrintResult(StringUtils.RightAfter(str, "CA", True) = "bcd", 7)

End Sub

Private Sub TestRotate()
    Dim str     As String

    Debug.Print "--- TestRotate ---"

    Call PrintResult(StringUtils.Rotate("", 0) = "", 1)
    Call PrintResult(StringUtils.Rotate("", 10) = "", 2)

    str = "abcdefg"
    Call PrintResult(StringUtils.Rotate(str, 0) = "abcdefg", 3)
    Call PrintResult(StringUtils.Rotate(str, 1) = "gabcdef", 4)
    Call PrintResult(StringUtils.Rotate(str, -1) = "bcdefga", 5)
    Call PrintResult(StringUtils.Rotate(str, 3) = "efgabcd", 6)
    Call PrintResult(StringUtils.Rotate(str, -4) = "efgabcd", 7)
    Call PrintResult(StringUtils.Rotate(str, 7) = "abcdefg", 8)
    Call PrintResult(StringUtils.Rotate(str, -7) = "abcdefg", 9)
    Call PrintResult(StringUtils.Rotate(str, 9) = "fgabcde", 10)
    Call PrintResult(StringUtils.Rotate(str, -13) = "gabcdef", 11)
    Call PrintResult(StringUtils.Rotate(str, 21) = "abcdefg", 12)
    Call PrintResult(StringUtils.Rotate(str, -28) = "abcdefg", 13)

End Sub

Private Sub TestRTrim()
    Dim str     As String

    Debug.Print "--- TestRTrim ---"

    str = ""
    Call PrintResult(StringUtils.RTrim(str) = "", 1)

    str = "ab c"
    Call PrintResult(StringUtils.RTrim(str) = "ab c", 2)

    str = " ab c "
    Call PrintResult(StringUtils.RTrim(str) = " ab c", 3)

    str = " ab c" & vbTab & vbCrLf & vbLf & " 　"
    Call PrintResult(StringUtils.RTrim(str) = " ab c", 4)

    str = "a "
    Call PrintResult(StringUtils.RTrim(str) = "a", 5)

    str = " 　"
    Call PrintResult(StringUtils.RTrim(str) = "", 6)

End Sub

Private Sub TestSplitByAlpha()
    Dim str     As String
    Dim arr()   As String

    Debug.Print "--- TestSplitByAlpha ---"

    arr = StringUtils.SplitByAlpha(str)
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)

    str = "a1b2c3d"
    arr = StringUtils.SplitByAlpha(str)
    Call PrintResult(LangUtils.ToString(arr) = "(1, 2, 3)", 2)

    arr = StringUtils.SplitByAlpha(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(a, 1, b, 2, c, 3, d)", 3)

    str = "012_a_345_AzByCxDwEvFuGtHsIrJqKpLoMnNmOlPkQjRiShTgUfVeWdXcYbZa_678"
    arr = StringUtils.SplitByAlpha(str)
    Call PrintResult(LangUtils.ToString(arr) = "(012_, _345_, _678)", 4)

    arr = StringUtils.SplitByAlpha(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(012_, a, _345_, AzByCxDwEvFuGtHsIrJqKpLoMnNmOlPkQjRiShTgUfVeWdXcYbZa, _678)", 5)

End Sub

Private Sub TestSplitByBlank()
    Dim str     As String
    Dim arr()   As String

    Debug.Print "--- TestSplitByBlank ---"

    arr = StringUtils.SplitByBlank(str)
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)

    str = "a " & vbCr & "b　" & vbLf & "c" & vbCrLf & "d"
    arr = StringUtils.SplitByBlank(str)
    Call PrintResult(LangUtils.ToString(arr) = "(a, b, c, d)", 2)

    arr = StringUtils.SplitByBlank(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(a,  " & vbCr & ", b, 　" & vbLf & ", c, " & vbCrLf & ", d)", 3)

    str = vbCrLf & "   " & vbCrLf & 123 & vbLf & "　　" & vbLf
    arr = StringUtils.SplitByBlank(str)
    Call PrintResult(LangUtils.ToString(arr) = "(123)", 4)

End Sub

Private Sub TestSplitByChars()
    Dim str     As String
    Dim arr()   As String

    Debug.Print "--- TestSplitByChars ---"

    arr = StringUtils.SplitByChars(str, "")
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)
    arr = StringUtils.SplitByChars(str, "abc")
    Call PrintResult(LangUtils.ToString(arr) = "()", 2)

    str = "_0_a+/1-b_2+c/3"
    arr = StringUtils.SplitByChars(str, "_")
    Call PrintResult(LangUtils.ToString(arr) = "(0, a+/1-b, 2+c/3)", 3)
    arr = StringUtils.SplitByChars(str, "_", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_, 0, _, a+/1-b, _, 2+c/3)", 4)

    arr = StringUtils.SplitByChars(str, "ab")
    Call PrintResult(LangUtils.ToString(arr) = "(_0_, +/1-, _2+c/3)", 5)
    arr = StringUtils.SplitByChars(str, "ab", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_0_, a, +/1-, b, _2+c/3)", 6)

    arr = StringUtils.SplitByChars(str, "123")
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+/, -b_, +c/)", 7)
    arr = StringUtils.SplitByChars(str, "123", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+/, 1, -b_, 2, +c/, 3)", 8)

    arr = StringUtils.SplitByChars(str, "_+-/")
    Call PrintResult(LangUtils.ToString(arr) = "(0, a, 1, b, 2, c, 3)", 9)
    arr = StringUtils.SplitByChars(str, "_+-/", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_, 0, _, a, +/, 1, -, b, _, 2, +, c, /, 3)", 10)

    arr = StringUtils.SplitByChars(str, "1b/")
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+, -, _2+c, 3)", 11)
    arr = StringUtils.SplitByChars(str, "1b/", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+, /1, -, b, _2+c, /, 3)", 12)

    arr = StringUtils.SplitByChars(str, "0123abc_+-/")
    Call PrintResult(LangUtils.ToString(arr) = "()", 13)
    arr = StringUtils.SplitByChars(str, "0123abc_+-/", True)
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+/1-b_2+c/3)", 14)

    arr = StringUtils.SplitByChars(str, "")
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+/1-b_2+c/3)", 15)
    arr = StringUtils.SplitByChars(str, "4")
    Call PrintResult(LangUtils.ToString(arr) = "(_0_a+/1-b_2+c/3)", 16)

End Sub

Private Sub TestSplitByDigit()
    Dim str     As String
    Dim arr()   As String

    Debug.Print "--- TestSplitByDigit ---"

    arr = StringUtils.SplitByDigit(str)
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)

    str = "0a1b2c3"
    arr = StringUtils.SplitByDigit(str)
    Call PrintResult(LangUtils.ToString(arr) = "(a, b, c)", 2)

    arr = StringUtils.SplitByDigit(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(0, a, 1, b, 2, c, 3)", 3)

    str = "abc_0_def_123456789_ghi"
    arr = StringUtils.SplitByDigit(str)
    Call PrintResult(LangUtils.ToString(arr) = "(abc_, _def_, _ghi)", 4)

    arr = StringUtils.SplitByDigit(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(abc_, 0, _def_, 123456789, _ghi)", 5)

End Sub

Private Sub TestSplitByNewline()
    Dim str     As String
    Dim arr()   As String

    Debug.Print "--- TestSplitByNewline ---"

    arr = StringUtils.SplitByNewline(str)
    Call PrintResult(LangUtils.ToString(arr) = "()", 1)

    str = "a " & vbCr & "b　" & vbLf & "c" & vbCrLf & "d"
    arr = StringUtils.SplitByNewline(str)
    Call PrintResult(LangUtils.ToString(arr) = "(a , b　, c, d)", 2)

    arr = StringUtils.SplitByNewline(str, True)
    Call PrintResult(LangUtils.ToString(arr) = "(a , " & vbCr & ", b　, " & vbLf & ", c, " & vbCrLf & ", d)", 3)

    str = vbCrLf & vbCrLf & vbCrLf & vbCrLf & 123 & vbLf & vbLf & vbLf & vbLf
    arr = StringUtils.SplitByNewline(str)
    Call PrintResult(LangUtils.ToString(arr) = "(123)", 4)

End Sub

Private Sub TestStartsWith()
    Dim str     As String

    Debug.Print "--- TestStartsWith ---"

    str = ""
    Call PrintResult(Not StringUtils.StartsWith(str, ""), 1)

    str = ""
    Call PrintResult(Not StringUtils.StartsWith(str, "a"), 2)

    str = "abcdef"
    Call PrintResult(StringUtils.StartsWith(str, "a"), 3)
    Call PrintResult(StringUtils.StartsWith(str, "abcdef"), 4)
    Call PrintResult(Not StringUtils.StartsWith(str, "b"), 5)
    Call PrintResult(Not StringUtils.StartsWith(str, "ABC"), 6)
    Call PrintResult(StringUtils.StartsWith(str, "ABC", True), 7)

End Sub

Private Sub TestStartsWithAny()
    Dim str         As String
    Dim params()    As String

    Debug.Print "--- TestStartsWithAny ---"

    str = ""
    Call PrintResult(Not StringUtils.StartsWithAny(str, ""), 1)
    Call PrintResult(Not StringUtils.StartsWithAny(str, params), 2)

    str = "abcdef"
    params = ArrayUtils.CStrArray("a", "ab", "abc")
    Call PrintResult(Not StringUtils.StartsWithAny(str, ""), 3)
    Call PrintResult(StringUtils.StartsWithAny(str, "abc"), 4)
    Call PrintResult(StringUtils.StartsWithAny(str, params), 5)
    Call PrintResult(StringUtils.StartsWithAny(str, "abc", "def", "1"), 6)
    Call PrintResult(Not StringUtils.StartsWithAny(str, "", "bcd", "1"), 7)

End Sub

Private Sub TestTrim()
    Dim str     As String

    Debug.Print "--- TestTrim ---"

    str = ""
    Call PrintResult(StringUtils.Trim(str) = "", 1)

    str = "ab c"
    Call PrintResult(StringUtils.Trim(str) = "ab c", 2)

    str = " ab c "
    Call PrintResult(StringUtils.Trim(str) = "ab c", 3)

    str = " 　ab c" & vbTab & vbCrLf & vbLf & " 　"
    Call PrintResult(StringUtils.Trim(str) = "ab c", 4)

    str = " a　"
    Call PrintResult(StringUtils.Trim(str) = "a", 5)

    str = " 　"
    Call PrintResult(StringUtils.Trim(str) = "", 6)

End Sub

