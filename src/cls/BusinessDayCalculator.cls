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
' MODULE   : �c�Ɠ������v�Z����N���X
'
'------------------------------------------------------------------------------

Private dayOffResolver_         As Object

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肳�ꂽ���t���x�������肷��I�u�W�F�N�g��ݒ肵�܂��B
'            �I�u�W�F�N�g�͉��L�̃��\�b�h����������Ă���K�v������܂��B
'
'            Function IsDayOff(ByRef d As Date) As Boolean
'
' PARAMS   : resolver - ����I�u�W�F�N�g
'
'------------------------------------------------------------------------------
Public Sub SetDayOffResolver(ByRef resolver As Object)

    Set dayOffResolver_ = resolver

End Sub

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肵��2�̓��t�̊Ԃ̉c�Ɠ������J�E���g���܂��B
'            �J�E���g���J�n������t�����͊܂܂�܂���B
'            �J�n���t���I�����t���ߋ��̏ꍇ�͕�����Ԃ��܂��B
'
' PARAMS   : fromDate - �J�n���t
'            toDate - �I�����t
'
' RETURN   : �c�Ɠ���
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
' FUNCTION : �w�肵�����t����w��c�Ɠ��������Z�������t���擾���܂��B
'            �c�Ɠ����ɕ������w�肷�邱�Ƃ��\�ł��B
'
' PARAMS   : d - ���t
'            offset - ���Z����c�Ɠ���
'
' RETURN   : �c�Ɠ��������Z�������t
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

    ' �y�����H
    If Weekday(d) = vbSaturday Or Weekday(d) = vbSunday Then Exit Function

    ' �����̏j���A�U�֋x���A�����̋x�����H
    If JapaneseHolidayUtils.IsHoliday(d) Then Exit Function

    IsDayOff = False

End Function