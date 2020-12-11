Attribute VB_Name = "JapaneseHolidayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE   : 日本の「国民の祝日」、「振替休日」、「国民の休日」に関するユーティリティクラス
'
'            [注意事項]
'            ・今後の法律改正により正常に動作しなくなる可能性があります。
'            ・2151年以降の「春分の日」、「秋分の日」は求めることができません。
'            ・複数の「国民の祝日」が重なる場合は、いずれか一方の情報しか取得できません。
'              （今後最初に祝日が重複するのは2876年と予測されています）
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付の「国民の祝日」の名前を取得します。
'            「振替休日」「国民の休日」の場合はそれぞれ"振替休日"、"国民の休日"を返します。
'            「国民の祝日」「振替休日」「国民の休日」のいずれでもない場合は空文字を返します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付の「国民の祝日」の名前、または"振替休日"、"国民の休日"
'
' ERROR    : 「春分の日」「秋分の日」が計算ができない年を指定された場合
'
'------------------------------------------------------------------------------
Public Function GetHolidayName(ByRef d As Date) As String
    Dim name        As String

    GetHolidayName = ""

    name = GetNationalHolidayName(d)
    If name <> "" Then
        GetHolidayName = name
        Exit Function
    End If

    ' 日曜の場合は、振替休日、国民の休日にはならない
    If Weekday(d) = vbSunday Then Exit Function

    name = GetSubstituteHoliday(d)
    If name <> "" Then
        GetHolidayName = name
        Exit Function
    End If

    GetHolidayName = GetCitizensHoliday(d)

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付の「国民の祝日」の名前を取得します。
'            「国民の祝日」ではない場合は空文字を返します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付の「国民の祝日」の名前
'
' ERROR    : 「春分の日」「秋分の日」が計算ができない年を指定された場合
'
'------------------------------------------------------------------------------
Public Function GetNationalHolidayName(ByRef d As Date) As String
    Dim y       As Integer
    Dim dy      As Integer

    GetNationalHolidayName = ""

    If DateDiff("d", "1948/7/20", d) < 0 Then Exit Function

    y = year(d)
    dy = Day(d)

    Select Case month(d)
        Case 1
            If dy = 1 Then
                GetNationalHolidayName = "元日"
                Exit Function
            End If

            If y <= 1999 Then
                If dy = 15 Then
                    GetNationalHolidayName = "成人の日"
                    Exit Function
                End If
            ElseIf 2000 <= y Then
                If dy = MondayOf(y, 1, 2) Then
                    GetNationalHolidayName = "成人の日"
                    Exit Function
                End If
            End If

        Case 2
            If 1967 <= y Then
                If dy = 11 Then
                    GetNationalHolidayName = "建国記念の日"
                    Exit Function
                End If
            End If

            If 2020 <= y Then
                If dy = 23 Then
                    GetNationalHolidayName = "天皇誕生日"
                    Exit Function
                End If
            End If

        Case 3
            If dy = CalcVernalEquinoxDay(y) Then
                GetNationalHolidayName = "春分の日"
                Exit Function
            End If

        Case 4
            If dy = 29 Then
                If y <= 1988 Then
                    GetNationalHolidayName = "天皇誕生日"
                ElseIf 1989 <= y And y <= 2006 Then
                    GetNationalHolidayName = "みどりの日"
                ElseIf 2007 <= y Then
                    GetNationalHolidayName = "昭和の日"
                End If
                Exit Function
            End If

        Case 5
            If 2007 <= y Then
                If dy = 4 Then
                    GetNationalHolidayName = "みどりの日"
                    Exit Function
                End If
            End If

            If dy = 3 Then
                GetNationalHolidayName = "憲法記念日"
                Exit Function
            End If

            If dy = 5 Then
                GetNationalHolidayName = "こどもの日"
                Exit Function
            End If

        Case 6

        Case 7
            If 1996 <= y And y <= 2002 Then
                If dy = 20 Then
                    GetNationalHolidayName = "海の日"
                    Exit Function
                End If
            ElseIf y = 2020 Then
                If dy = 23 Then
                    GetNationalHolidayName = "海の日"
                    Exit Function
                End If
            ElseIf 2003 <= y Then
                If dy = MondayOf(y, 7, 3) Then
                    GetNationalHolidayName = "海の日"
                    Exit Function
                End If
            End If

            If y = 2020 Then
                If dy = 24 Then
                    GetNationalHolidayName = "スポーツの日"
                    Exit Function
                End If
            End If

        Case 8
            If y = 2020 Then
                If dy = 10 Then
                    GetNationalHolidayName = "山の日"
                    Exit Function
                End If
            ElseIf 2016 <= y Then
                If dy = 11 Then
                    GetNationalHolidayName = "山の日"
                    Exit Function
                End If
            End If

        Case 9
            If 1966 <= y And y <= 2002 Then
                If dy = 15 Then
                    GetNationalHolidayName = "敬老の日"
                    Exit Function
                End If
            ElseIf 2003 <= y Then
                If dy = MondayOf(y, 9, 3) Then
                    GetNationalHolidayName = "敬老の日"
                    Exit Function
                End If
            End If

            If dy = CalcAutumnalEquinoxDay(y) Then
                GetNationalHolidayName = "秋分の日"
                Exit Function
            End If

        Case 10
            If 1966 <= y And y <= 1999 Then
                If dy = 10 Then
                    GetNationalHolidayName = "体育の日"
                    Exit Function
                End If
            ElseIf 2000 <= y And y <= 2019 Then
                If dy = MondayOf(y, 10, 2) Then
                    GetNationalHolidayName = "体育の日"
                    Exit Function
                End If
            ElseIf 2021 <= y Then
                If dy = MondayOf(y, 10, 2) Then
                    GetNationalHolidayName = "スポーツの日"
                    Exit Function
                End If
            End If

        Case 11
            If dy = 3 Then
                GetNationalHolidayName = "文化の日"
                Exit Function
            End If

            If dy = 23 Then
                GetNationalHolidayName = "勤労感謝の日"
                Exit Function
            End If

        Case 12
            If 1989 <= y And y <= 2018 Then
                If dy = 23 Then
                    GetNationalHolidayName = "天皇誕生日"
                    Exit Function
                End If
            End If
    End Select

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付が「国民の祝日」「振替休日」「国民の休日」のいずれかであるか判定します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付が「国民の祝日」「振替休日」「国民の休日」のいずれかである場合は true
'
' ERROR    : 「春分の日」「秋分の日」が計算ができない年を指定された場合
'
'------------------------------------------------------------------------------
Public Function IsHoliday(ByRef d As Date) As Boolean

    If GetHolidayName(d) <> "" Then
        IsHoliday = True
    Else
        IsHoliday = False
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付が「国民の祝日」であるか判定します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付が「国民の祝日」である場合は true
'
' ERROR    : 「春分の日」「秋分の日」が計算ができない年を指定された場合
'
'------------------------------------------------------------------------------
Public Function IsNationalHoliday(ByRef d As Date) As Boolean

    If GetNationalHolidayName(d) <> "" Then
        IsNationalHoliday = True
    Else
        IsNationalHoliday = False
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付が「振替休日」であるか判定します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付が「振替休日」である場合は true
'
'------------------------------------------------------------------------------
Public Function IsSubstituteHoliday(ByRef d As Date) As Boolean

    If GetHolidayName(d) = "振替休日" Then
        IsSubstituteHoliday = True
    Else
        IsSubstituteHoliday = False
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定した日付が「国民の休日」であるか判定します。
'
' PARAMS   : d - 日付
'
' RETURN   : 指定した日付が「国民の休日」である場合は true
'
'------------------------------------------------------------------------------
Public Function IsCitizensHoliday(ByRef d As Date) As Boolean

    If GetHolidayName(d) = "国民の休日" Then
        IsCitizensHoliday = True
    Else
        IsCitizensHoliday = False
    End If

End Function

Private Function MondayOf(ByVal year As Integer, ByVal month As Integer, ByVal ordinal As Integer) As Integer
    Dim wd      As Integer

    wd = Weekday(DateSerial(year, month, 1))
    If wd <= vbMonday Then
        MondayOf = vbMonday - wd + 1 + 7 * (ordinal - 1)
    Else
        MondayOf = vbMonday - wd + 1 + 7 * ordinal
    End If

End Function

Private Function CalcVernalEquinoxDay(ByVal year As Integer) As Integer
    Dim diff        As Integer
    Dim standard    As Double

    diff = year - 1980

    If year <= 1979 Then
        standard = 20.8357
    ElseIf year <= 2099 Then
        standard = 20.8431
    ElseIf year <= 2150 Then
        standard = 21.851
    Else
        Call Err.Raise(5)
    End If

    CalcVernalEquinoxDay = Int(standard + 0.242194 * diff - Int(diff / 4))

End Function

Private Function CalcAutumnalEquinoxDay(ByVal year As Integer) As Integer
    Dim diff        As Integer
    Dim standard    As Double

    diff = year - 1980

    If year <= 1979 Then
        standard = 23.2588
    ElseIf year <= 2099 Then
        standard = 23.2488
    ElseIf year <= 2150 Then
        standard = 24.2488
    Else
        Call Err.Raise(5)
    End If

    CalcAutumnalEquinoxDay = Int(standard + 0.242194 * diff - Int(diff / 4))

End Function

Private Function GetSubstituteHoliday(ByRef d As Date) As String
    Dim tmpD        As Date

    GetSubstituteHoliday = ""

    ' 法律改正前は、振替休日にはならない
    If DateDiff("d", "1973/4/12", d) < 0 Then Exit Function

    tmpD = DateAdd("d", -1, d)
    If DateDiff("d", "2007/1/1", d) < 0 Then
        ' 祝日が日曜日の場合はその翌日の月曜日を振替休日とする
        If IsNationalHoliday(tmpD) And Weekday(tmpD) = vbSunday Then
            GetSubstituteHoliday = "振替休日"
        End If
    Else
        ' 連続する祝日のうち、どれか1日が日曜日と重なった場合は、最後の祝日の翌日が振替休日とする
        Do While True
            If IsNationalHoliday(tmpD) Then
                If Weekday(tmpD) = vbSunday Then
                    GetSubstituteHoliday = "振替休日"
                End If
            Else
                Exit Do
            End If
            tmpD = DateAdd("d", -1, tmpD)
        Loop
    End If

End Function

Private Function GetCitizensHoliday(ByRef d As Date) As String

    GetCitizensHoliday = ""

    ' 法律改正前は、国民の休日にはならない
    If DateDiff("d", "1985/12/27", d) < 0 Then Exit Function

    If IsNationalHoliday(DateAdd("d", -1, d)) And _
        IsNationalHoliday(DateAdd("d", 1, d)) Then
        GetCitizensHoliday = "国民の休日"
    End If

End Function
