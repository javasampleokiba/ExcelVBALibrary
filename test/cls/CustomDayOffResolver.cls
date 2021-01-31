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
' MODULE   : BusinessDayCalculator�e�X�g���W���[���p�̃N���X
'
'------------------------------------------------------------------------------

Function IsDayOff(ByRef d As Date) As Boolean

    IsDayOff = True

    ' �y�����H
    If Weekday(d) = vbSaturday Or Weekday(d) = vbSunday Then Exit Function

    ' �����̏j���A�U�֋x���A�����̋x�����H
    If JapaneseHolidayUtils.IsHoliday(d) Then Exit Function

    ' �O�����͋x��
    If month(d) = 1 And (Day(d) = 2 Or Day(d) = 3) Then Exit Function

    IsDayOff = False

End Function