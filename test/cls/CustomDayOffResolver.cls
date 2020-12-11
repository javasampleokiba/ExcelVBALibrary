VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomDayOffResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE   : BusinessDayCalculatorテストモジュール用のクラス
'
'------------------------------------------------------------------------------

Function IsDayOff(ByRef d As Date) As Boolean

    IsDayOff = True

    ' 土日か？
    If Weekday(d) = vbSaturday Or Weekday(d) = vbSunday Then Exit Function

    ' 国民の祝日、振替休日、国民の休日か？
    If JapaneseHolidayUtils.IsHoliday(d) Then Exit Function

    ' 三が日は休日
    If month(d) = 1 And (Day(d) = 2 Or Day(d) = 3) Then Exit Function

    IsDayOff = False

End Function
