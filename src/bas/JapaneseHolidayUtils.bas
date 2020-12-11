Attribute VB_Name = "JapaneseHolidayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE   : ���{�́u�����̏j���v�A�u�U�֋x���v�A�u�����̋x���v�Ɋւ��郆�[�e�B���e�B�N���X
'
'            [���ӎ���]
'            �E����̖@�������ɂ�萳��ɓ��삵�Ȃ��Ȃ�\��������܂��B
'            �E2151�N�ȍ~�́u�t���̓��v�A�u�H���̓��v�͋��߂邱�Ƃ��ł��܂���B
'            �E�����́u�����̏j���v���d�Ȃ�ꍇ�́A�����ꂩ����̏�񂵂��擾�ł��܂���B
'              �i����ŏ��ɏj�����d������̂�2876�N�Ɨ\������Ă��܂��j
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肵�����t�́u�����̏j���v�̖��O���擾���܂��B
'            �u�U�֋x���v�u�����̋x���v�̏ꍇ�͂��ꂼ��"�U�֋x��"�A"�����̋x��"��Ԃ��܂��B
'            �u�����̏j���v�u�U�֋x���v�u�����̋x���v�̂�����ł��Ȃ��ꍇ�͋󕶎���Ԃ��܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t�́u�����̏j���v�̖��O�A�܂���"�U�֋x��"�A"�����̋x��"
'
' ERROR    : �u�t���̓��v�u�H���̓��v���v�Z���ł��Ȃ��N���w�肳�ꂽ�ꍇ
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

    ' ���j�̏ꍇ�́A�U�֋x���A�����̋x���ɂ͂Ȃ�Ȃ�
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
' FUNCTION : �w�肵�����t�́u�����̏j���v�̖��O���擾���܂��B
'            �u�����̏j���v�ł͂Ȃ��ꍇ�͋󕶎���Ԃ��܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t�́u�����̏j���v�̖��O
'
' ERROR    : �u�t���̓��v�u�H���̓��v���v�Z���ł��Ȃ��N���w�肳�ꂽ�ꍇ
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
                GetNationalHolidayName = "����"
                Exit Function
            End If

            If y <= 1999 Then
                If dy = 15 Then
                    GetNationalHolidayName = "���l�̓�"
                    Exit Function
                End If
            ElseIf 2000 <= y Then
                If dy = MondayOf(y, 1, 2) Then
                    GetNationalHolidayName = "���l�̓�"
                    Exit Function
                End If
            End If

        Case 2
            If 1967 <= y Then
                If dy = 11 Then
                    GetNationalHolidayName = "�����L�O�̓�"
                    Exit Function
                End If
            End If

            If 2020 <= y Then
                If dy = 23 Then
                    GetNationalHolidayName = "�V�c�a����"
                    Exit Function
                End If
            End If

        Case 3
            If dy = CalcVernalEquinoxDay(y) Then
                GetNationalHolidayName = "�t���̓�"
                Exit Function
            End If

        Case 4
            If dy = 29 Then
                If y <= 1988 Then
                    GetNationalHolidayName = "�V�c�a����"
                ElseIf 1989 <= y And y <= 2006 Then
                    GetNationalHolidayName = "�݂ǂ�̓�"
                ElseIf 2007 <= y Then
                    GetNationalHolidayName = "���a�̓�"
                End If
                Exit Function
            End If

        Case 5
            If 2007 <= y Then
                If dy = 4 Then
                    GetNationalHolidayName = "�݂ǂ�̓�"
                    Exit Function
                End If
            End If

            If dy = 3 Then
                GetNationalHolidayName = "���@�L�O��"
                Exit Function
            End If

            If dy = 5 Then
                GetNationalHolidayName = "���ǂ��̓�"
                Exit Function
            End If

        Case 6

        Case 7
            If 1996 <= y And y <= 2002 Then
                If dy = 20 Then
                    GetNationalHolidayName = "�C�̓�"
                    Exit Function
                End If
            ElseIf y = 2020 Then
                If dy = 23 Then
                    GetNationalHolidayName = "�C�̓�"
                    Exit Function
                End If
            ElseIf 2003 <= y Then
                If dy = MondayOf(y, 7, 3) Then
                    GetNationalHolidayName = "�C�̓�"
                    Exit Function
                End If
            End If

            If y = 2020 Then
                If dy = 24 Then
                    GetNationalHolidayName = "�X�|�[�c�̓�"
                    Exit Function
                End If
            End If

        Case 8
            If y = 2020 Then
                If dy = 10 Then
                    GetNationalHolidayName = "�R�̓�"
                    Exit Function
                End If
            ElseIf 2016 <= y Then
                If dy = 11 Then
                    GetNationalHolidayName = "�R�̓�"
                    Exit Function
                End If
            End If

        Case 9
            If 1966 <= y And y <= 2002 Then
                If dy = 15 Then
                    GetNationalHolidayName = "�h�V�̓�"
                    Exit Function
                End If
            ElseIf 2003 <= y Then
                If dy = MondayOf(y, 9, 3) Then
                    GetNationalHolidayName = "�h�V�̓�"
                    Exit Function
                End If
            End If

            If dy = CalcAutumnalEquinoxDay(y) Then
                GetNationalHolidayName = "�H���̓�"
                Exit Function
            End If

        Case 10
            If 1966 <= y And y <= 1999 Then
                If dy = 10 Then
                    GetNationalHolidayName = "�̈�̓�"
                    Exit Function
                End If
            ElseIf 2000 <= y And y <= 2019 Then
                If dy = MondayOf(y, 10, 2) Then
                    GetNationalHolidayName = "�̈�̓�"
                    Exit Function
                End If
            ElseIf 2021 <= y Then
                If dy = MondayOf(y, 10, 2) Then
                    GetNationalHolidayName = "�X�|�[�c�̓�"
                    Exit Function
                End If
            End If

        Case 11
            If dy = 3 Then
                GetNationalHolidayName = "�����̓�"
                Exit Function
            End If

            If dy = 23 Then
                GetNationalHolidayName = "�ΘJ���ӂ̓�"
                Exit Function
            End If

        Case 12
            If 1989 <= y And y <= 2018 Then
                If dy = 23 Then
                    GetNationalHolidayName = "�V�c�a����"
                    Exit Function
                End If
            End If
    End Select

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肵�����t���u�����̏j���v�u�U�֋x���v�u�����̋x���v�̂����ꂩ�ł��邩���肵�܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t���u�����̏j���v�u�U�֋x���v�u�����̋x���v�̂����ꂩ�ł���ꍇ�� true
'
' ERROR    : �u�t���̓��v�u�H���̓��v���v�Z���ł��Ȃ��N���w�肳�ꂽ�ꍇ
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
' FUNCTION : �w�肵�����t���u�����̏j���v�ł��邩���肵�܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t���u�����̏j���v�ł���ꍇ�� true
'
' ERROR    : �u�t���̓��v�u�H���̓��v���v�Z���ł��Ȃ��N���w�肳�ꂽ�ꍇ
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
' FUNCTION : �w�肵�����t���u�U�֋x���v�ł��邩���肵�܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t���u�U�֋x���v�ł���ꍇ�� true
'
'------------------------------------------------------------------------------
Public Function IsSubstituteHoliday(ByRef d As Date) As Boolean

    If GetHolidayName(d) = "�U�֋x��" Then
        IsSubstituteHoliday = True
    Else
        IsSubstituteHoliday = False
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肵�����t���u�����̋x���v�ł��邩���肵�܂��B
'
' PARAMS   : d - ���t
'
' RETURN   : �w�肵�����t���u�����̋x���v�ł���ꍇ�� true
'
'------------------------------------------------------------------------------
Public Function IsCitizensHoliday(ByRef d As Date) As Boolean

    If GetHolidayName(d) = "�����̋x��" Then
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

    ' �@�������O�́A�U�֋x���ɂ͂Ȃ�Ȃ�
    If DateDiff("d", "1973/4/12", d) < 0 Then Exit Function

    tmpD = DateAdd("d", -1, d)
    If DateDiff("d", "2007/1/1", d) < 0 Then
        ' �j�������j���̏ꍇ�͂��̗����̌��j����U�֋x���Ƃ���
        If IsNationalHoliday(tmpD) And Weekday(tmpD) = vbSunday Then
            GetSubstituteHoliday = "�U�֋x��"
        End If
    Else
        ' �A������j���̂����A�ǂꂩ1�������j���Əd�Ȃ����ꍇ�́A�Ō�̏j���̗������U�֋x���Ƃ���
        Do While True
            If IsNationalHoliday(tmpD) Then
                If Weekday(tmpD) = vbSunday Then
                    GetSubstituteHoliday = "�U�֋x��"
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

    ' �@�������O�́A�����̋x���ɂ͂Ȃ�Ȃ�
    If DateDiff("d", "1985/12/27", d) < 0 Then Exit Function

    If IsNationalHoliday(DateAdd("d", -1, d)) And _
        IsNationalHoliday(DateAdd("d", 1, d)) Then
        GetCitizensHoliday = "�����̋x��"
    End If

End Function
