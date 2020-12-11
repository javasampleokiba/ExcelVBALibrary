Attribute VB_Name = "TestSuite"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : 全モジュールのテストを実行するモジュール
'
'------------------------------------------------------------------------------

Sub TestAll()
    Dim val1            As Byte
    Dim val2            As Boolean
    Dim val3            As Integer
    Dim val4            As Long
    Dim val5            As Single
    Dim val6            As Double
    Dim val7            As Currency
    Dim val8            As Object
    Dim val9            As String
    Dim val10           As String * 20
    Dim val11(1)        As String
    Dim arr(1 To 18)    As Variant

    ' 全データ型を格納したテスト用配列を作成
    val1 = 10
    val2 = True
    val3 = 100
    val4 = 100000
    val5 = 0.001
    val6 = 0.00000000001
    val7 = 1000
    Set val8 = ActiveSheet
    val9 = "AAA"
    val10 = "BBB"
    val11(0) = "CCC"
    val11(1) = "DDD"

    arr(1) = val1       ' Byte
    arr(2) = val2       ' Boolean
    arr(3) = val3       ' Integer
    arr(4) = val4       ' Long
    arr(5) = val5       ' Single
    arr(6) = val6       ' Double
    arr(7) = val7       ' Currency
    Set arr(8) = val8   ' Object
    arr(9) = val9       ' String
    arr(10) = val10     ' String * 20
    arr(11) = val11     ' String配列
    arr(12) = Empty
    arr(13) = Null
    arr(14) = Err
    arr(15) = vbNullString
    arr(16) = vbNullChar
    arr(17) = vbNull
    Set arr(18) = Nothing

    ' [標準モジュール]
    Call TestArrayUtils.TestAll(arr)
    Call TestCellAddressUtils.TestAll
    Call TestJapaneseHolidayUtils.TestAll
    Call TestLangUtils.TestAll(arr)

    ' [クラスモジュール]
    Call TestBusinessDayCalculator.TestAll
    
End Sub

Public Sub PrintResult(ByVal result As Boolean, Optional ByVal num As Integer = 0)

    If result Then
        If num = 0 Then
            Debug.Print "OK!"
        Else
            Debug.Print "No." & num & " OK!"
        End If
    Else
        If num = 0 Then
            Debug.Print "NG!"
        Else
            Debug.Print "No." & num & " NG!"
        End If
    End If

End Sub

Public Sub PrintResultIfNg(ByVal result As Boolean, Optional ByVal num As Integer = 0)

    If Not result Then
        If num = 0 Then
            Debug.Print "NG!"
        Else
            Debug.Print "No." & num & " NG!"
        End If
    End If

End Sub
