Attribute VB_Name = "TestBusinessDayCalculator"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : BusinessDayCalculator�̃e�X�g���W���[��
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

    ' �J�n���t = �I�����t ������
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/1/2020") = 0)
    ' �J�n���t = �I�����t ���y�j
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/4/2020") = 0)
    ' �J�n���t = �I�����t �����j
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/5/2020") = 0)
    ' �J�n�I�����t������Ⴂ ������
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/2/2020") = 1)
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "3/31/2020") = -1)
    ' �J�n�I�����t������Ⴂ ���J�n���y�j
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/5/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/3/2020") = -1)
    ' �J�n�I�����t������Ⴂ ���J�n�����j
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/6/2020") = 1)
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/4/2020") = 0)
    ' �J�n�I�����t������Ⴂ ���I�����y�j
    Call PrintResultIfNg(bdc.CountDays("4/3/2020", "4/4/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/5/2020", "4/4/2020") = 0)
    ' �J�n�I�����t������Ⴂ ���I�������j
    Call PrintResultIfNg(bdc.CountDays("4/4/2020", "4/5/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("4/6/2020", "4/5/2020") = 0)
    ' �����̂�
    Call PrintResultIfNg(bdc.CountDays("4/6/2020", "4/10/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("4/9/2020", "4/7/2020") = -2)
    ' �T���܂����j���Ȃ� ���J�n�I��������
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "4/15/2020") = 10)
    Call PrintResultIfNg(bdc.CountDays("4/28/2020", "4/17/2020") = -7)
    ' �T���܂����j���Ȃ� ���J�n��not����
    Call PrintResultIfNg(bdc.CountDays("4/11/2020", "4/22/2020") = 8)
    Call PrintResultIfNg(bdc.CountDays("4/26/2020", "4/13/2020") = -10)
    ' �T���܂����j���Ȃ� ���I����not����
    Call PrintResultIfNg(bdc.CountDays("4/15/2020", "4/19/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("4/13/2020", "4/5/2020") = -5)
    ' �j���̂�
    Call PrintResultIfNg(bdc.CountDays("5/3/2020", "5/6/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("7/24/2020", "7/23/2020") = 0)
    ' �����Əj���̂� ���J�n�I��������
    Call PrintResultIfNg(bdc.CountDays("4/27/2020", "4/30/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("11/6/2020", "11/2/2020") = -3)
    ' �����Əj���̂� ���J�n���j��
    Call PrintResultIfNg(bdc.CountDays("8/10/2020", "8/14/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("9/25/2020", "9/21/2020") = -2)
    ' �����Əj���̂� ���I�����j��
    Call PrintResultIfNg(bdc.CountDays("11/2/2020", "11/3/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("11/27/2020", "11/23/2020") = -3)
    ' �����Əj���̂� ���J�n�I�����j��
    Call PrintResultIfNg(bdc.CountDays("9/20/2021", "9/23/2021") = 2)
    Call PrintResultIfNg(bdc.CountDays("5/6/2022", "5/2/2022") = -1)
    ' �y���Əj���̂� ����
    Call PrintResultIfNg(bdc.CountDays("5/2/2020", "5/6/2020") = 0)
    Call PrintResultIfNg(bdc.CountDays("11/23/2020", "11/21/2020") = 0)
    ' �����A�y���A�j�����ׂĊ܂� ����T�ԓ�
    Call PrintResultIfNg(bdc.CountDays("2/9/2020", "2/15/2020") = 4)
    Call PrintResultIfNg(bdc.CountDays("7/25/2020", "7/19/2020") = -3)
    ' �����A�y���A�j�����ׂĊ܂� �������܂���
    Call PrintResultIfNg(bdc.CountDays("4/1/2020", "5/31/2020") = 38)
    Call PrintResultIfNg(bdc.CountDays("9/30/2020", "7/1/2020") = -60)
    ' �����A�y���A�j�����ׂĊ܂� ���N���܂���
    Call PrintResultIfNg(bdc.CountDays("11/1/2020", "1/31/2021") = 61)
    Call PrintResultIfNg(bdc.CountDays("2/29/2020", "11/1/2019") = -81)
    ' �U�֋x���A�����̋x�����܂�
    Call PrintResultIfNg(bdc.CountDays("5/1/2020", "5/8/2020") = 2)
    Call PrintResultIfNg(bdc.CountDays("9/25/2015", "9/21/2015") = -1)
    ' �����̏j������O�̂����N�̉c�Ɠ���
    Call PrintResultIfNg(bdc.CountDays("1/1/1947", "12/31/1947") = 260)
    ' �����̏j�������̂����N�̉c�Ɠ���
    Call PrintResultIfNg(bdc.CountDays("1/1/2000", "12/31/2000") = 249)

    Set resolver = New CustomDayOffResolver
    Call bdc.SetDayOffResolver(resolver)

    ' �O�����͋x���Ƃ���p�^�[��
    Call PrintResultIfNg(bdc.CountDays("1/1/2020", "1/10/2020") = 5)

End Sub

Private Sub TestGetDate()
    Dim bdc         As BusinessDayCalculator
    Dim resolver    As CustomDayOffResolver

    Debug.Print "--- TestGetDate ---"

    Set bdc = New BusinessDayCalculator

    ' �c�Ɠ���=0 ������
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 0) = "4/1/2020")
    ' �c�Ɠ���=0 ����not����
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", 0) = "4/4/2020")
    ' �c�Ɠ���=1 �������̂�
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 1) = "4/2/2020")
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", -1) = "3/31/2020")
    ' �c�Ɠ���=1 ���y�����܂���
    Call PrintResultIfNg(bdc.GetDate("4/3/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/5/2020", 1) = "4/6/2020")
    Call PrintResultIfNg(bdc.GetDate("4/6/2020", -1) = "4/3/2020")
    Call PrintResultIfNg(bdc.GetDate("4/5/2020", -1) = "4/3/2020")
    Call PrintResultIfNg(bdc.GetDate("4/4/2020", -1) = "4/3/2020")
    ' �c�Ɠ���=1 ���j�����܂���
    Call PrintResultIfNg(bdc.GetDate("4/28/2020", 1) = "4/30/2020")
    Call PrintResultIfNg(bdc.GetDate("1/2/2020", -1) = "12/31/2019")
    ' �c�Ɠ���=2�ȏ� �������̂�
    Call PrintResultIfNg(bdc.GetDate("4/6/2020", 4) = "4/10/2020")
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", -2) = "3/30/2020")
    ' �c�Ɠ���=2�ȏ� ���y�����܂���
    Call PrintResultIfNg(bdc.GetDate("4/1/2020", 5) = "4/8/2020")
    Call PrintResultIfNg(bdc.GetDate("4/13/2020", -6) = "4/3/2020")
    ' �c�Ɠ���=2�ȏ� ���j�����܂���
    Call PrintResultIfNg(bdc.GetDate("2/9/2020", 4) = "2/14/2020")
    Call PrintResultIfNg(bdc.GetDate("5/1/2020", -3) = "4/27/2020")
    ' �c�Ɠ���=2�ȏ� ���c�Ɠ����A�������ɓy���j�����܂���
    Call PrintResultIfNg(bdc.GetDate("2/7/2020", 11) = "2/26/2020")
    Call PrintResultIfNg(bdc.GetDate("8/12/2020", -12) = "7/22/2020")
    Call PrintResultIfNg(bdc.GetDate("5/5/2020", 4) = "5/12/2020")
    Call PrintResultIfNg(bdc.GetDate("9/22/2020", -8) = "9/9/2020")
    ' �����̏j������O�̂����N�̉c�Ɠ���
    Call PrintResultIfNg(bdc.GetDate("1/1/1947", 260) = "12/31/1947")
    ' �����̏j�������̂����N�̉c�Ɠ���
    Call PrintResultIfNg(bdc.GetDate("1/1/2000", 249) = "12/29/2000")

    Set resolver = New CustomDayOffResolver
    Call bdc.SetDayOffResolver(resolver)

    ' �O�����͋x���Ƃ���p�^�[��
    Call PrintResultIfNg(bdc.GetDate("1/1/2020", 5) = "1/10/2020")

End Sub
