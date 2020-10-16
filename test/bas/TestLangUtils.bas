Attribute VB_Name = "TestLangUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : LangUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll(ByVal arr As Variant)

    Debug.Print "=== TestLangUtils ==="
    
    Call TestIsEqual
    Call TestIsEqual2(arr)
    Call TestToString(arr)

End Sub

Private Sub TestIsEqual()
    Dim i           As Integer
    Dim j           As Integer
    Dim a1(12)      As Variant
    Dim a2()        As Variant
    Dim header(12)  As String
    Dim report      As String
    Dim result      As Boolean
    Dim cur         As Currency
    Dim arr(1)      As String
    Dim arr2(1)     As Variant
    Dim arr1(1)     As String
    
    Debug.Print "--- TestIsEqual ---"
    
    report = "<html><head><title>Report</title></head><body><table border=1>"
    
    cur = 1
    arr(1) = "AAA"
    arr1(0) = "BBB"
    arr1(1) = "CCC"
    arr2(1) = arr1
    
    header(0) = "Integer"
    header(1) = "String"
    header(2) = "Null"
    header(3) = "Empty"
    header(4) = "Nothing"
    header(5) = "Error"
    header(6) = "Date"
    header(7) = "Currency"
    header(8) = "vbNullString"
    header(9) = "vbNullChar"
    header(10) = "Object"
    header(11) = "Array"
    header(12) = "2D Array"
    
    a1(0) = 1
    a1(1) = "1"
    a1(2) = Null
    a1(3) = Empty
    Set a1(4) = Nothing
    a1(5) = Error
    a1(6) = Date
    a1(7) = cur
    a1(8) = vbNullString
    a1(9) = vbNullChar
    Set a1(10) = ActiveSheet
    a1(11) = arr
    a1(12) = arr2
    a2 = a1
    
    report = report & "<tr>"
    report = report & "<td></td>"
    For i = 0 To UBound(header)
        report = report & "<td>" & header(i) & "</td>"
    Next
    report = report & "</tr>"
    For i = 0 To UBound(a1)
        report = report & "<tr>"
        report = report & "<td>" & header(i) & "</td>"
        For j = 0 To UBound(a2)
            result = LangUtils.IsEqual(a1(i), a2(j))
            If result Then
                report = report & "<td style=""color: red;"">" & result & "</td>"
            Else
                report = report & "<td>" & result & "</td>"
            End If
        Next j
        report = report & "</tr>"
    Next i
    
    report = report & "</table></body></html>"
    Call PrintReport(report)
    
End Sub

Private Sub TestIsEqual2(ByVal arr As Variant)
    Dim a1()        As Integer
    Dim a2()        As Integer
    Dim a3()        As Integer
    Dim a4()        As Integer
    
    Debug.Print "--- TestIsEqual2 ---"
    
    ReDim a2(0 To 0)
    ReDim a3(1 To 1)
    
    ' 全データ型で比較ができるか確認
    Call PrintResult(LangUtils.IsEqual(arr, arr), 1)

    ' 両方未初期化で確認
    Call PrintResult(LangUtils.IsEqual(a1, a4), 2)

    ' 片方未初期化で確認
    Call PrintResult(Not LangUtils.IsEqual(a1, a2), 3)

    ' サイズは同じだが範囲が違うもの同士を確認
    Call PrintResult(Not LangUtils.IsEqual(a2, a3), 4)
    
End Sub

Private Sub TestToString(ByVal arr As Variant)
    Dim arr2()          As Variant
    Dim myClazz         As MyClass

    Set myClazz = New MyClass
    
    Debug.Print "--- TestToString ---"
    
    ' 全データ型が出力できるか確認
    Debug.Print LangUtils.ToString(arr)

    ' 空の配列
    Debug.Print LangUtils.ToString(arr2)

    ' Nothingの場合
    Debug.Print LangUtils.ToString(Nothing)

    ' 標準のオブジェクト型の場合
    Debug.Print LangUtils.ToString(ActiveSheet)

    ' 自作クラスの場合
    Debug.Print LangUtils.ToString(myClazz)

End Sub

Private Sub PrintReport(ByVal data As String)
    Dim fno As Integer
    fno = FreeFile
    Open ThisWorkbook.Path & "/report.html" For Output As #fno

    Print #fno, data

    Close #fno
    
    Debug.Print "report.htmlを出力しました。"

End Sub
