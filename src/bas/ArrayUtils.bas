Attribute VB_Name = "ArrayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : �z�񑀍�Ɋւ��郆�[�e�B���e�B���W���[��
'
' LINK   : ���̃��W���[���͈ȉ��̃��W���[�����Q�Ƃ��Ă��܂��B
'
'          �ELangUtils
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肳�ꂽ�z�񂪋�ł��邩���肵�܂��B
'            ����������Ă��Ȃ��A���邢��Erase���s��̓��I�z��̏ꍇ��
'            true ��Ԃ��܂��B
'
' PARAMS   : arr - ����Ώۂ̔z��
'
' RETURN   : �w�肳�ꂽ�z�񂪋�ł���ꍇ�� true
'
' ERROR    : �����ɔz��ȊO���w�肵���ꍇ
'
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef arr As Variant) As Boolean
    Dim i       As Long

    If Not IsArray(arr) Then
        Call Err.Raise(5)   ' �v���V�[�W���̌Ăяo���A�܂��͈������s���ł��B
    End If
    
    On Error GoTo ErrHandler
    i = UBound(arr)
    IsEmptyArray = False

    Exit Function

ErrHandler:
    IsEmptyArray = True

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肳�ꂽ2�̔z�񂪓����������肵�܂��B
'
' PARAMS   : arr1 - ��r�Ώۂ̔z��1
'            arr2 - ��r�Ώۂ̔z��2
'
' RETURN   : �w�肳�ꂽ2�̔z�񂪓������ꍇ�� true
'
' ERROR    : �����ɔz��ȊO���w�肵���ꍇ
'
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    Dim i       As Long

    IsEqual = False

    ' �����Ƃ��v�f�����ݒ肳��Ă��Ȃ����I�z��̏ꍇ
    If IsEmptyArray(arr1) And IsEmptyArray(arr2) Then
        IsEqual = True
        Exit Function
        
    ' �����ꂩ���v�f�����ݒ肳��Ă��Ȃ����I�z��̏ꍇ
    ElseIf IsEmptyArray(arr1) Or IsEmptyArray(arr2) Then
        Exit Function
    End If

    ' �z��̍ŏ��E�ő�C���f�b�N�X���������Ȃ��ꍇ
    If LBound(arr1) <> LBound(arr2) Then Exit Function
    If UBound(arr1) <> UBound(arr2) Then Exit Function

    ' �S�v�f���r
    For i = LBound(arr1) To UBound(arr1)
        If Not LangUtils.IsEqual(arr1(i), arr2(i)) Then
            Exit Function
        End If
    Next

    IsEqual = True

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �z��̗v�f�����擾���܂��B
'
' PARAMS   : arr - �擾�Ώۂ̔z��
'
' RETURN   : �z��̗v�f��
'
'------------------------------------------------------------------------------
Public Function Length(ByRef arr As Variant) As Long

    If IsEmptyArray(arr) Then
        Length = 0
        Exit Function
    End If

    Length = UBound(arr) - LBound(arr) + 1

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肳�ꂽ�z��̗v�f�������Ƀ\�[�g���܂��B
'            ���ׂĂ̗v�f���s�������Z�q�ɂ���r���\�ł��邱�Ƃ��O������ł��B
'            CompareTo���\�b�h�����������I�u�W�F�N�g�̔z��������ɂł��܂��B
'
' PARAMS   : arr - �����Ώۂ̔z��
'
'------------------------------------------------------------------------------
Public Sub Sort(ByRef arr As Variant)

    If Length(arr) <= 1 Then
        Exit Sub
    End If

    If IsObject(arr(LBound(arr))) Then
        Call QuickObjSort(arr, LBound(arr), UBound(arr))
    Else
        Call QuickSort(arr, LBound(arr), UBound(arr))
    End If

End Sub

Private Sub QuickSort(ByRef arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i       As Long
    Dim j       As Long
    Dim pivot   As Variant
    Dim tmp     As Variant

    If left >= right Then
        Exit Sub
    End If

    i = left
    j = right
    pivot = arr((i + j) / 2)

    Do
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While pivot < arr(j)
            j = j - 1
        Loop
        If j <= i Then Exit Do
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        i = i + 1
        j = j - 1
    Loop

    Call QuickSort(arr, left, i - 1)
    Call QuickSort(arr, j + 1, right)

End Sub

Private Sub QuickObjSort(ByRef arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i       As Long
    Dim j       As Long
    Dim pivot   As Variant
    Dim tmp     As Variant

    If left >= right Then
        Exit Sub
    End If

    i = left
    j = right
    Set pivot = arr((i + j) / 2)

    Do
        Do While 0 < pivot.CompareTo(arr(i))
            i = i + 1
        Loop
        Do While pivot.CompareTo(arr(j)) < 0
            j = j - 1
        Loop
        If j <= i Then Exit Do
        Set tmp = arr(i)
        Set arr(i) = arr(j)
        Set arr(j) = tmp
        i = i + 1
        j = j - 1
    Loop

    Call QuickObjSort(arr, left, i - 1)
    Call QuickObjSort(arr, j + 1, right)

End Sub
