Attribute VB_Name = "TestBusinessDayCalculator"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : BusinessDayCalculatorのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll()

    Debug.Print "=== TestBusinessDayCalculator ==="
    
    Call TestCountDays
    Call TestGetDate

End Sub

Private Sub TestCountDays()
    Dim bdc         As BusinessDayCalculator
    Dim resolver    As CustomDayOffResolver

    Debug.Print "--- TestCountDays ---"

    Set bdc = New BusinessDayCalculator

    ' 開始日付 = 終了日付 かつ平日
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/1/2020") = 0)
    ' 開始日付 = 終了日付 かつ土曜
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/4/2020") = 0)
    ' 開始日付 = 終了日付 かつ日曜
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/5/2020") = 0)
    ' 開始終了日付が一日違い かつ平日
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/2/2020") = 1)
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "3/31/2020") = -1)
    ' 開始終了日付が一日違い かつ開始が土曜
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/5/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/3/2020") = -1)
    ' 開始終了日付が一日違い かつ開始が日曜
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/6/2020") = 1)
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/4/2020") = 0)
    ' 開始終了日付が一日違い かつ終了が土曜
    Call PrintResultIfNg(bdc.CountDays("4/3/2020", "4/4/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/4/2020") = 0)
    ' 開始終了日付が一日違い かつ終了が日曜
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/5/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/6/2020", "4/5/2020") = 0)
    ' 平日のみ
    Call PrintResultIfNg(bdc.CountDays("4/6/2020", "4/10/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("4/9/2020", "4/7/2020") = -2)
    ' 週をまたぎ祝日なし かつ開始終了が平日
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/15/2020") = 10)
    Call PrintResultIfNg(bdc.CountDays("4/28/2020", "4/17/2020") = -7)
    ' 週をまたぎ祝日なし かつ開始がnot平日
    Call PrintResultIfNg(bdc.CountDays("4/11/2020", "4/22/2020") = 8)
    Call PrintResultIfNg(bdc.CountDays("4/26/2020", "4/13/2020") = -10)
    ' 週をまたぎ祝日なし かつ終了がnot平日
    Call PrintResultIfNg(bdc.CountDays("4/15/2020", "4/19/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("4/13/2020", "4/5/2020") = -5)
    ' 祝日のみ
    Call PrintResultIfNg(bdc.CountDays("5/3/2020", "5/6/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("7/24/2020", "7/23/2020") = 0)
    ' 平日と祝日のみ かつ開始終了が平日
    Call PrintResultIfNg(bdc.CountDays("4/27/2020", "4/30/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("11/6/2020", "11/2/2020") = -3)
    ' 平日と祝日のみ かつ開始が祝日
    Call PrintResultIfNg(bdc.CountDays("8/10/2020", "8/14/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("9/25/2020", "9/21/2020") = -2)
    ' 平日と祝日のみ かつ終了が祝日
    Call PrintResultIfNg(bdc.CountDays("11/2/2020", "11/3/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("11/27/2020", "11/23/2020") = -3)
    ' 平日と祝日のみ かつ開始終了が祝日
    Call PrintResultIfNg(bdc.CountDays("9/20/2021", "9/23/2021") = 2)
    Call PrintResultIfNg(bdc.CountDays("5/6/2022", "5/2/2022") = -1)
    ' 土日と祝日のみ かつ
    Call PrintResultIfNg(bdc.CountDays("5/2/2020", "5/6/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("11/23/2020", "11/21/2020") = 0)
    ' 平日、土日、祝日すべて含む かつ一週間内
    Call PrintResultIfNg(bdc.CountDays("2/9/2020", "2/15/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("7/25/2020", "7/19/2020") = -3)
    ' 平日、土日、祝日すべて含む かつ月をまたぐ
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "5/31/2020") = 38)
    Call PrintResultIfNg(bdc.CountDays("9/30/2020", "7/1/2020") = -60)
    ' 平日、土日、祝日すべて含む かつ年をまたぐ
    Call PrintResultIfNg(bdc.CountDays("11/1/2020", "1/31/2021") = 61)
    Call PrintResultIfNg(bdc.CountDays("2/29/2020", "11/1/2019") = -81)
    ' 振替休日、国民の休日を含む
    Call PrintResultIfNg(bdc.CountDays("5/1/2020", "5/8/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("9/25/2015", "9/21/2015") = -1)
    ' 国民の祝日制定前のある一年の営業日数
    Call PrintResultIfNg(bdc.CountDays("1/1/1947", "12/31/1947") = 260)
    ' 国民の祝日制定後のある一年の営業日数
    Call PrintResultIfNg(bdc.CountDays("1/1/2000", "12/31/2000") = 249)

    Set resolver = New CustomDayOffResolver
    Call bdc.SetDayOffResolver(resolver)

    ' 三が日は休日とするパターン
    Call PrintResultIfNg(bdc.CountDays("1/1/2020", "1/10/2020") = 5)

End Sub

Private Sub TestGetDate()
    Dim bdc         As BusinessDayCalculator
    Dim resolver    As CustomDayOffResolver

    Debug.Print "--- TestGetDate ---"

    Set bdc = New BusinessDayCalculator

    ' 営業日数=0 かつ平日
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 0) = "4/1/2020")
    ' 営業日数=0 かつnot平日
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", 0) = "4/4/2020")
    ' 営業日数=1 かつ平日のみ
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 1) = "4/2/2020")
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", -1) = "3/31/2020")
    ' 営業日数=1 かつ土日をまたぐ
    Call PrintResultIfNg(bdc.GetDate("4/3/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/5/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/6/2020", -1) = "4/3/2020")
    Call PrintResultIfNg(bdc.GetDate("4/5/2020", -1) = "4/3/2020")
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", -1) = "4/3/2020")
    ' 営業日数=1 かつ祝日をまたぐ
    Call PrintResultIfNg(bdc.GetDate("4/28/2020", 1) = "4/30/2020")
    Call PrintResultIfNg(bdc.GetDate("1/2/2020", -1) = "12/31/2019")
    ' 営業日数=2以上 かつ平日のみ
    Call PrintResultIfNg(bdc.GetDate("4/6/2020", 4) = "4/10/2020")
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", -2) = "3/30/2020")
    ' 営業日数=2以上 かつ土日をまたぐ
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 5) = "4/8/2020")
    Call PrintResultIfNg(bdc.GetDate("4/13/2020", -6) = "4/3/2020")
    ' 営業日数=2以上 かつ祝日をまたぐ
    Call PrintResultIfNg(bdc.GetDate("2/9/2020", 4) = "2/14/2020")
    Call PrintResultIfNg(bdc.GetDate("5/1/2020", -3) = "4/27/2020")
    ' 営業日数=2以上 かつ営業日が連続せずに土日祝日をまたぐ
    Call PrintResultIfNg(bdc.GetDate("2/7/2020", 11) = "2/26/2020")
    Call PrintResultIfNg(bdc.GetDate("8/12/2020", -12) = "7/22/2020")
    Call PrintResultIfNg(bdc.GetDate("5/5/2020", 4) = "5/12/2020")
    Call PrintResultIfNg(bdc.GetDate("9/22/2020", -8) = "9/9/2020")
    ' 国民の祝日制定前のある一年の営業日数
    Call PrintResultIfNg(bdc.GetDate("1/1/1947", 260) = "12/31/1947")
    ' 国民の祝日制定後のある一年の営業日数
    Call PrintResultIfNg(bdc.GetDate("1/1/2000", 249) = "12/29/2000")

    Set resolver = New CustomDayOffResolver
    Call bdc.SetDayOffResolver(resolver)

    ' 三が日は休日とするパターン
    Call PrintResultIfNg(bdc.GetDate("1/1/2020", 5) = "1/10/2020")

End Sub
