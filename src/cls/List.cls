VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' MODULE : �����t����ꂽ�����̗v�f���i�[����R���N�V�����N���X
'
' NOTE   : �f�t�H���g�ł͂ǂ�ȃf�[�^�^�ł����X�g�Ɋi�[�ł��܂��B
'          �^���S���̂��ߊi�[����f�[�^�^���w�肵�����ꍇ�́A
'          ���̃I�u�W�F�N�g���������SetDataType���\�b�h�����s���Ă��������B
'
' LINK   : ���̃��W���[���͈ȉ��̃��W���[�����Q�Ƃ��Ă��܂��B
'
'          �EArrayUtils
'          �ELangUtils
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' �����o�ϐ�
'------------------------------------------------------------------------------
Private dataType_       As String
Private arr_()          As Variant
Private returnItem_     As Variant

'------------------------------------------------------------------------------
' �C���X�^���X����������
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

    dataType_ = ""

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g���̎w��ʒu�Ɏw�肳�ꂽ�v�f��ǉ����܂��B
'            �ʒu�C���f�b�N�X�ɕ������w�肷��ƃ��X�g�̌�������
'            �ʒu�C���f�b�N�X�ɒǉ����܂��B
'
' PARAMS   : index - �ǉ�����ʒu�C���f�b�N�X
'            item  - �ǉ�����v�f
'
' ERROR    : �ʒu�C���f�b�N�X���z��͈̔͊O�̏ꍇ
'------------------------------------------------------------------------------
Public Sub Add(ByVal index As Long, ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Add(arr_, index, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g���炷�ׂĂ̗v�f���폜���܂��B
'------------------------------------------------------------------------------
Public Sub Clear()

    Erase arr_

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̃R�s�[�𐶐����ĕԂ��܂��B
'
' RETURN   : �R�s�[�������X�g
'------------------------------------------------------------------------------
Public Function Clone() As List
    Dim l   As List

    Set l = New List
    Call l.Concat(arr_)
    Set Clone = l

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�Ɏw�肳�ꂽ�C�ӂ̐��̗v�f��ǉ����܂��B
'            �v�f�����X�g�̏ꍇ�́A���̃��X�g�Ɋ܂܂�邷�ׂĂ̗v�f��ǉ����܂��B
'
' PARAMS   : items - �ǉ�����C�ӂ̐��̗v�f
'------------------------------------------------------------------------------
Public Sub Concat(ParamArray items() As Variant)
    Dim flattenArr()    As Variant
    Dim arr()           As Variant
    Dim item            As Variant

    flattenArr = ArrayUtils.Flatten(2, items)

    If IsEmptyArray(flattenArr) Then
        Exit Sub
    End If

    For Each item In flattenArr
        If IsList(item) Then
            Call ArrayUtils.Concat(arr, item.ToArray())
        Else
            Call ArrayUtils.Concat(arr, Array(item))
        End If
    Next

    Call CheckDataType(arr)

    Call ArrayUtils.Concat(arr_, arr)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���v�f�����̃��X�g�Ɋ܂܂�邩���肵�܂��B
'
' PARAMS   : item - ��������v�f
'
' RETURN   : ���݂���ꍇ�� True
'------------------------------------------------------------------------------
Public Function Contains(ByRef item As Variant) As Boolean

    Contains = ArrayUtils.Contains(arr_, item)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���C�ӂ̐��̗v�f���ׂĂ����̃��X�g�Ɋ܂܂�邩���肵�܂��B
'
' PARAMS   : items - ��������C�ӂ̐��̗v�f
'
' RETURN   : ���ׂĂ̗v�f�����݂���ꍇ�� True
'------------------------------------------------------------------------------
Public Function ContainsAll(ParamArray items() As Variant) As Boolean

    ContainsAll = ArrayUtils.ContainsAll(arr_, ArrayUtils.Flatten(2, items))

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���C�ӂ̐��̗v�f�̂����ꂩ�����̃��X�g�Ɋ܂܂�邩���肵�܂��B
'
' PARAMS   : items - ��������C�ӂ̐��̗v�f
'
' RETURN   : �����ꂩ�̗v�f�����݂���ꍇ�� True
'------------------------------------------------------------------------------
Public Function ContainsAny(ParamArray items() As Variant) As Boolean

    ContainsAny = ArrayUtils.ContainsAny(arr_, ArrayUtils.Flatten(2, items))

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肳�ꂽ�v�f�����̃��X�g�ɂ������݂��邩�J�E���g���܂��B
'
' PARAMS   : item - ��������v�f
'
' RETURN   : ���������v�f�̌�
'------------------------------------------------------------------------------
Public Function Count(ByRef item As Variant) As Long

    Count = ArrayUtils.Count(arr_, item)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̐擪�̗v�f���擾���܂��B
'
' RETURN   : �擪�̗v�f
'------------------------------------------------------------------------------
Public Function First() As Variant

    If IsObjectType(ArrayUtils.First(arr_)) Then
        Set First = returnItem_
    Else
        First = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����w�肵���ʒu�C���f�b�N�X�̗v�f���擾���܂��B
'            �ʒu�C���f�b�N�X�ɕ������w�肷��ƃ��X�g�̌�������
'            �ʒu�C���f�b�N�X�̗v�f���擾���܂��B
'
' PARAMS   : index - �擾����ʒu�C���f�b�N�X
'            [def] - �ʒu�C���f�b�N�X���s���̏ꍇ�ɕԂ��f�t�H���g�l (�ȗ�����Null)
'
' RETURN   : �w�肵���ʒu�C���f�b�N�X�̗v�f �܂��� �f�t�H���g�l
'------------------------------------------------------------------------------
Public Function GetAt(ByVal index As Long, Optional ByRef def As Variant = Null) As Variant

    If IsObjectType(ArrayUtils.GetAt(arr_, index, def)) Then
        Set GetAt = returnItem_
    Else
        GetAt = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�Ɋi�[�ł���f�[�^�^�𕶎���Ŏ擾���܂��B
'
' RETURN   : ���̃��X�g�Ɋi�[�ł���f�[�^�^
'------------------------------------------------------------------------------
Public Function GetDataType() As String

    GetDataType = dataType_

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���v�f�����̃��X�g���ōŏ��Ɍ��������ʒu�C���f�b�N�X��
'            �Ԃ��܂��B
'
' PARAMS   : item    - ��������v�f
'            [start] - �����J�n�ʒu�C���f�b�N�X (�͈͊O�ł����������삵�܂�)
'
' RETURN   : �w�肵���v�f�����������ʒu�C���f�b�N�X (���݂��Ȃ��ꍇ�� -1)
'------------------------------------------------------------------------------
Public Function IndexOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long

    IndexOf = ArrayUtils.IndexOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���v�f�����̃��X�g���Ō����������ׂĂ̈ʒu�C���f�b�N�X��
'            �Ԃ��܂��B
'
' PARAMS   : item    - ��������v�f
'            [start] - �����J�n�ʒu�C���f�b�N�X (�͈͊O�ł����������삵�܂�)
'
' RETURN   : �w�肵���v�f�����������ʒu�C���f�b�N�X�̔z��
'            (���݂��Ȃ��ꍇ�͋�̔z��)
'------------------------------------------------------------------------------
Public Function IndicesOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long()

    IndicesOf = ArrayUtils.IndicesOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�ɗv�f���Ȃ��ꍇ��True��Ԃ��܂��B
'
' RETURN   : ���̃��X�g�ɗv�f���Ȃ��ꍇ��True
'------------------------------------------------------------------------------
Public Function IsEmptyList() As Boolean

    IsEmptyList = ArrayUtils.IsEmptyArray(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�Ǝw�肳�ꂽ���X�g�������������肵�܂��B
'
' PARAMS   : other - ��r�Ώۂ̃��X�g
'
' RETURN   : �w�肳�ꂽ���X�g�Ɠ������ꍇ�� True
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef other As List) As Boolean

    IsEqual = ArrayUtils.IsEqual(arr_, other.ToArray())

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̊e�v�f�̕�����\�����w�肳�ꂽ��؂蕶���ŘA�����A
'            �Ԃ��܂��B
'
' PARAMS   : [sep] - ��؂蕶�� (�ȗ����ꂽ�ꍇ�͋󕶎�)
'
' RETURN   : �A�����ꂽ������
'------------------------------------------------------------------------------
Public Function Join(Optional ByVal sep As String = "") As String

    Join = ArrayUtils.Join(arr_, sep)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̍Ō�̗v�f���擾���܂��B
'
' RETURN   : �Ō�̗v�f
'------------------------------------------------------------------------------
Public Function Last() As Variant

    If IsObjectType(ArrayUtils.Last(arr_)) Then
        Set Last = returnItem_
    Else
        Last = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : �w�肵���v�f�����̃��X�g���ōŌ�Ɍ��������ʒu�C���f�b�N�X��
'            �Ԃ��܂��B
'
' PARAMS   : item    - ��������v�f
'            [start] - �����J�n�ʒu�C���f�b�N�X (�͈͊O�ł����������삵�܂�)
'
' RETURN   : �w�肵���v�f�����������ʒu�C���f�b�N�X (���݂��Ȃ��ꍇ�� -1)
'------------------------------------------------------------------------------
Public Function LastIndexOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long

    LastIndexOf = ArrayUtils.LastIndexOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : �Ōォ�猟�����āA�w�肵���v�f�����̃��X�g���Ō�������
'            ���ׂĂ̈ʒu�C���f�b�N�X��Ԃ��܂��B
'
' PARAMS   : item    - ��������v�f
'            [start] - �����J�n�ʒu�C���f�b�N�X (�͈͊O�ł����������삵�܂�)
'
' RETURN   : �w�肵���v�f�����������ʒu�C���f�b�N�X�̔z��
'            (���݂��Ȃ��ꍇ�͋�̔z��)
'------------------------------------------------------------------------------
Public Function LastIndicesOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long()

    LastIndicesOf = ArrayUtils.LastIndicesOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g���ɂ���v�f�̐���Ԃ��܂��B
'
' RETURN   : �v�f�̐�
'------------------------------------------------------------------------------
Public Function Length() As Long

    Length = ArrayUtils.Length(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̒��ōő�̗v�f���擾���܂��B
'            ���ׂĂ̗v�f���s�������Z�q�ɂ���r���\�ł��邱�Ƃ��O������ł��B
'
' RETURN   : �ő�̗v�f
'------------------------------------------------------------------------------
Public Function Max() As Variant

    If IsObjectType(ArrayUtils.Max(arr_)) Then
        Set Max = returnItem_
    Else
        Max = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̒��ōŏ��̗v�f���擾���܂��B
'            ���ׂĂ̗v�f���s�������Z�q�ɂ���r���\�ł��邱�Ƃ��O������ł��B
'
' RETURN   : �ŏ��̗v�f
'------------------------------------------------------------------------------
Public Function Min() As Variant

    If IsObjectType(ArrayUtils.Min(arr_)) Then
        Set Min = returnItem_
    Else
        Min = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̍Ō�̗v�f���폜���ĕԂ��܂��B
'
' RETURN   : �폜�����v�f
'------------------------------------------------------------------------------
Public Function Pop() As Variant

    If IsObjectType(ArrayUtils.Pop(arr_)) Then
        Set Pop = returnItem_
    Else
        Pop = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̍Ō�ɗv�f��ǉ����܂��B
'
' PARAMS   : item - �ǉ�����v�f
'------------------------------------------------------------------------------
Public Sub Push(ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Push(arr_, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����w�肵���v�f���폜���܂��B
'
' PARAMS   : item   - �폜����v�f
'            [size] - �폜����ő吔 (�ȗ�����ƁA���ׂč폜)
'------------------------------------------------------------------------------
Public Sub Remove(ByRef item As Variant, Optional ByVal size As Long = 0)

    Call ArrayUtils.Remove(arr_, item, size)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����w�肵���C�ӂ̐��̗v�f���ׂĂ��폜���܂��B
'
' PARAMS   : items  - �폜����C�ӂ̐��̗v�f
'------------------------------------------------------------------------------
Public Sub RemoveAll(ParamArray items() As Variant)

    Call ArrayUtils.RemoveAll(arr_, items)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g���̎w�肵���ʒu�̗v�f���폜���܂��B
'            �ʒu�C���f�b�N�X�ɕ������w�肷��ƃ��X�g�̌�������
'            �ʒu�C���f�b�N�X�̗v�f���폜���܂��B
'
' PARAMS   : index - �폜����ʒu�C���f�b�N�X
'
' RETURN   : �폜���ꂽ�v�f
'
' ERROR    : �ʒu�C���f�b�N�X���z��͈̔͊O�̏ꍇ
'------------------------------------------------------------------------------
Public Function RemoveAt(ByVal index As Long) As Variant

    If IsObjectType(ArrayUtils.RemoveAt(arr_, index)) Then
        Set RemoveAt = returnItem_
    Else
        RemoveAt = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̐擪����w�肵���v�f���������A�w�肵���ʂ̗v�f��
'            �u�������܂��B
'
' PARAMS   : item        - �u���Ώۗv�f
'            replacement - �u������v�f
'            [size]      - �u������ő吔 (�ȗ�����ƁA���ׂĒu��)
'
' RETURN   : �u�������v�f��
'------------------------------------------------------------------------------
Public Function Replace(ByRef item As Variant, ByRef replacement As Variant, _
                        Optional ByVal size As Long = 0) As Long

    Call CheckDataType(replacement)

    Replace = ArrayUtils.Replace(arr_, item, replacement, size)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̖�������w�肵���v�f���������A�w�肵���ʂ̗v�f�ɒu�������܂��B
'
' PARAMS   : item        - �u���Ώۗv�f
'            replacement - �u������v�f
'            [size]      - �u������ő吔 (�ȗ�����ƁA���ׂĒu��)
'
' RETURN   : �u�������v�f��
'------------------------------------------------------------------------------
Public Function ReplaceLast(ByRef item As Variant, ByRef replacement As Variant, _
                            Optional ByVal size As Long = 0) As Long

    Call CheckDataType(replacement)

    ReplaceLast = ArrayUtils.ReplaceLast(arr_, item, replacement, size)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f�̏��Ԃ𔽓]���܂��B
'------------------------------------------------------------------------------
Public Sub Reverse()

    Call ArrayUtils.Reverse(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f���w�肳�ꂽ����������]�����܂��B
'            �����ɐ��̒l���w�肷��Ɨv�f�̓��X�g�̌������Ɉړ����A
'            ���̒l���w�肷��ƑO�����Ɉړ����܂��B
'
' PARAMS   : distance - �v�f�̈ړ�����
'------------------------------------------------------------------------------
Public Sub Rotate(ByVal distance As Long)

    Call ArrayUtils.Rotate(arr_, distance)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f���w��������_���i�[�������X�g��Ԃ��܂��B
'
' PARAMS   : [size] - �擾����� (�ȗ������1�B�ő�͔z��T�C�Y)
'            [uniq] - �����ʒu�̗v�f�𕡐��擾�s�Ƃ��邩�̃t���O
'
' RETURN   : �I�΂ꂽ�v�f���琬��V�������X�g
'------------------------------------------------------------------------------
Public Function Sample(Optional ByVal size As Long = 1, _
                       Optional ByVal uniq As Boolean = True) As List
    Dim arr()   As Variant

    arr = ArrayUtils.Sample(arr_, size, uniq)
    Set Sample = NewList(arr)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̎w�肵���ʒu�C���f�b�N�X�ɗv�f��ݒ肵�܂��B
'            �ʒu�C���f�b�N�X�ɕ������w�肷��ƃ��X�g�̌�������
'            �ʒu�C���f�b�N�X�ɒǉ����܂��B
'
' PARAMS   : index - �ݒ肷��ʒu�C���f�b�N�X
'            item  - �ݒ肷��v�f
'
' RETURN   : �ݒ�O�Ɏw�肵���ʒu�C���f�b�N�X�ɂ������v�f
'------------------------------------------------------------------------------
Public Function SetAt(ByVal index As Long, ByRef item As Variant) As Variant

    Call CheckDataType(item)

    If IsObjectType(ArrayUtils.SetAt(arr_, index, item)) Then
        Set SetAt = returnItem_
    Else
        SetAt = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�Ɋi�[�ł���f�[�^�^�𕶎���Ŏw�肵�܂��B
'
' PARAMS   : dataType - �f�[�^�^
'
' RETURN   : ���X�g����̏ꍇ�͏��True��Ԃ��܂��B
'            ���łɃf�[�^���i�[����Ă���A�w�肳�ꂽ�f�[�^�^��
'            ��v���Ȃ��ꍇ��False��Ԃ��܂��B
'------------------------------------------------------------------------------
Public Function SetDataType(ByVal dataType As String) As Boolean
    Dim item        As Variant

    ' �z�񂪋�ł͂Ȃ��ꍇ�͐ݒ肷��O�ɐ�Ƀ`�F�b�N���s��
    If dataType <> "" And Not Me.IsEmptyList() Then
        For Each item In arr_
            If typeName(item) <> dataType Then
                SetDataType = False
                Exit Function
            End If
        Next
    End If

    dataType_ = dataType
    SetDataType = True

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̐擪�̗v�f���폜���ĕԂ��܂��B
'
' RETURN   : �폜�����v�f
'------------------------------------------------------------------------------
Public Function Shift() As Variant

    If IsObjectType(ArrayUtils.Shift(arr_)) Then
        Set Shift = returnItem_
    Else
        Shift = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f�̏��Ԃ������_���ɓ���ւ��܂��B
'------------------------------------------------------------------------------
Public Sub Shuffle()

    Call ArrayUtils.Shuffle(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g��������Ŏw�肵���͈̗͂v�f�����o���A
'            �V�������X�g�𐶐����ĕԂ��܂��B
'
' PARAMS   : [startIdx] - �擾����J�n�ʒu�C���f�b�N�X (�ȗ�����Ɛ擪)
'            [endIdx]   - �擾����I���ʒu�C���f�b�N�X�B�������AendIdx�̈ʒu��
'                         �܂܂�Ȃ��B (�ȗ�����Ɩ����܂Ŏ擾)
'
' RETURN   : ���o�����v�f���琬��V�������X�g
'------------------------------------------------------------------------------
Public Function Slice(Optional ByVal startIdx As Long = 0, _
                       Optional ByVal endIdx As Long = 0) As List

    Set Slice = NewList(ArrayUtils.Slice(arr_, startIdx, endIdx))

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f�������Ƀ\�[�g���܂��B
'            ���ׂĂ̗v�f���s�������Z�q�ɂ���r���\�ł��邱�Ƃ��O������ł��B
'------------------------------------------------------------------------------
Public Sub Sort()

    Call ArrayUtils.Sort(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����v�f����菜���A�V�����v�f��ǉ����܂��B
'            �ǉ�����v�f���w�肵�Ȃ��ꍇ�͒P�ɗv�f���폜���܂��B
'            �폜����0�ɂ���ƒP�ɗv�f��}�����܂��B
'
' PARAMS   : index     - ���X�g��ω������n�߂�C���f�b�N�X
'                        (�������w�肷��ƃ��X�g�̌������̈ʒu�C���f�b�N�X)
'            [size]    - �폜����ő�v�f�� (�ȗ�����ƊJ�n�ʒu������̑S�v�f���폜)
'            [addList] - �ǉ�����v�f���i�[�������X�g
'
' ERROR    : �ʒu�C���f�b�N�X�����X�g�͈̔͊O�̏ꍇ
'------------------------------------------------------------------------------
Public Sub Splice(ByVal index As Long, Optional ByVal size As Long = -1, _
                  Optional ByRef addList As List = Nothing)
    Dim arr()   As Variant

    If addList Is Nothing Then
        Call ArrayUtils.Splice(arr_, index, size)
    Else
        arr = addList.ToArray()
    
        Call CheckDataType(arr)

        Call ArrayUtils.Splice(arr_, index, size, arr)
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����2�̗v�f�̈ʒu�����ւ��܂��B
'
' PARAMS   : index1 - �ύX����ʒu�C���f�b�N�X1
'            index2 - �ύX����ʒu�C���f�b�N�X2
'------------------------------------------------------------------------------
Public Sub Swap(ByVal index1 As Long, ByVal index2 As Long)

    Call ArrayUtils.Swap(arr_, index1, index2)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̗v�f�ō\�������V�����z��𐶐����ĕԂ��܂��B
'
' RETURN   : ���̃��X�g�̗v�f�ō\�������V�����z��
'------------------------------------------------------------------------------
Public Function ToArray() As Variant()

    ToArray = arr_

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̕�����\����Ԃ��܂��B
'
' RETURN   : ���̃��X�g�̕�����\��
'------------------------------------------------------------------------------
Public Function ToString() As String

    ToString = LangUtils.ToString(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g����d������v�f�����ׂč폜���܂��B
'------------------------------------------------------------------------------
Public Sub Unique()

    Call ArrayUtils.Unique(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g�̐擪�ɗv�f��ǉ����܂��B
'
' PARAMS   : item - �ǉ�����v�f
'------------------------------------------------------------------------------
Public Sub Unshift(ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Unshift(arr_, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : ���̃��X�g��������Ŏw�肵���ʒu�̗v�f�����o���A
'            �V�������X�g�𐶐����ĕԂ��܂��B
'
' PARAMS   : indices - �擾����ʒu�C���f�b�N�X�A�܂��͈ʒu�C���f�b�N�X�̔z��
'
' RETURN   : ���o�����v�f���琬��V�������X�g
'------------------------------------------------------------------------------
Public Function ValuesAt(ParamArray indices() As Variant) As List
    Dim flattenArr()    As Variant
    Dim result()        As Variant

    flattenArr = ArrayUtils.Flatten(0, indices)
    result = ArrayUtils.ValuesAt(arr_, flattenArr)
    Set ValuesAt = NewList(result)

End Function

Private Sub CheckDataType(ByRef item As Variant)
    Dim var     As Variant

    If dataType_ <> "" Then
        If IsArray(item) Then
            If Not IsEmptyArray(item) Then
                For Each var In item
                    If typeName(var) <> dataType_ Then
                        Call Err.Raise(13)
                    End If
                Next
            End If
        Else
            If typeName(item) <> dataType_ Then
                Call Err.Raise(13)
            End If
        End If
    End If

End Sub

Private Function IsList(ByRef item As Variant) As Boolean

    If typeName(item) = "List" Then
        IsList = True
    Else
        IsList = False
    End If

End Function

Private Function IsObjectType(ByRef item As Variant) As Boolean

    If IsObject(item) Then
        Set returnItem_ = item
        IsObjectType = True
    Else
        returnItem_ = item
        IsObjectType = False
    End If

End Function

Private Function NewList(ByRef arr() As Variant) As List
    Dim l       As List

    Set l = New List
    Call l.Concat(arr)
    Set NewList = l

End Function
