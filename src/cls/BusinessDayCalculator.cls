VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BusinessDayCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE   : 営業日数を計算するクラス
'
'------------------------------------------------------------------------------

Private dayOffResolver_         As Object

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された日付が休日か判定するオブジェクトを設定します。
'            オブジェクトは下記のメソッドが実装されている必要があります。
'
'            Function IsDayOff(ByRef d As Date) As Boolean
'
' PARAMS   : resolver - 判定オブジェクト
'
'------------------------------------------------------------------------------
Public Sub SetDayOffResolver(ByRef resolver As Object)

    Set dayOffResolver_ = resolver

End Sub

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した2つの日付の間の営業日数をカウントします。
'            カウントを開始する日付当日は含まれません。
'            開始日付より終了日付が過去の場合は負数を返します。
'
' PARAMS   : fromDate - 開始日付
'            toDate - 終了日付
'
' RETURN   : 営業日数
'
'------------------------------------------------------------------------------
Public Function CountDays(ByRef fromDate As Date, ByRef toDate As Date) As Long
    Dim count       As Long
    Dim step        As Integer
    Dim tmpD        As Date

    CountDays = 0

    If DateDiff("d", fromDate, toDate) = 0 Then Exit Function

    tmpD = fromDate
    If DateDiff("d", fromDate, toDate) > 0 Then
        step = 1
    Else
        step = -1
    End If

    Do
        tmpD = DateAdd("d", step, tmpD)
        If dayOffResolver_ Is Nothing Then
            If Not IsDayOff(tmpD) Then count = count + 1
        Else
            If Not dayOffResolver_.IsDayOff(tmpD) Then count = count + 1
        End If

    Loop While DateDiff("d", tmpD, toDate) <> 0

    CountDays = count * step

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付から指定営業日数を加算した日付を取得します。
'            営業日数に負数を指定することも可能です。
'
' PARAMS   : d - 日付
'            offset - 加算する営業日数
'
' RETURN   : 営業日数を加算した日付
'
'------------------------------------------------------------------------------
Public Function GetDate(ByRef d As Date, ByVal offset As Long) As Date
    Dim tmpD        As Date
    Dim count       As Long
    Dim step        As Integer

    tmpD = d

    If offset = 0 Then
        GetDate = tmpD
        Exit Function
    End If

    count = 0
    If offset > 0 Then
        step = 1
    Else
        step = -1
    End If

    Do
        tmpD = DateAdd("d", step, tmpD)
        If dayOffResolver_ Is Nothing Then
            If Not IsDayOff(tmpD) Then count = count + 1
        Else
            If Not dayOffResolver_.IsDayOff(tmpD) Then count = count + 1
        End If

    Loop While count <> offset * step

    GetDate = tmpD

End Function

Private Function IsDayOff(ByRef d As Date) As Boolean

    IsDayOff = True

    ' 土日か？
    If Weekday(d) = vbSaturday Or Weekday(d) = vbSunday Then Exit Function

    ' 国民の祝日、振替休日、国民の休日か？
    If JapaneseHolidayUtils.IsHoliday(d) Then Exit Function

    IsDayOff = False

End Function
