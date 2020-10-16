Attribute VB_Name = "LangUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : ExcelVBA�̋��ʓI�����Ɏg�p����郆�[�e�B���e�B���W���[��
'
' LINK   : ���̃��W���[���͈ȉ��̃��W���[�����Q�Ƃ��Ă��܂��B
'
'          �EArrayUtils
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : 2�̈���������̒l�ł��邩���肵�܂��B
'            �z��̏ꍇ�͔z��̊e�v�f���r���A���ׂē���̒l�ł���� true
'            ��Ԃ��܂��B
'
' PARAMS   : item1 - ��r�Ώۂ̒l1
'            item2 - ��r�Ώۂ̒l2
'
' RETURN   : 2�̈���������̒l�ł���ꍇ�� true
'
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef item1 As Variant, ByRef item2 As Variant) As Boolean

    IsEqual = False

    ' �ǂ��炩����݂̂� Empty �̏ꍇ �˕s��v
    If IsEmpty(item1) Xor IsEmpty(item2) Then Exit Function

    ' �ǂ��炩����݂̂� Null �̏ꍇ �˕s��v
    If IsNull(item1) Xor IsNull(item2) Then Exit Function

    ' �ǂ��炩����݂̂� �z�� �̏ꍇ �˕s��v
    If IsArray(item1) Xor IsArray(item2) Then Exit Function

    ' �ǂ��炩����݂̂� Object�^ �̏ꍇ �˕s��v
    If IsObject(item1) Xor IsObject(item2) Then Exit Function

    ' �����Ƃ� Empty �̏ꍇ �ˈ�v
    If IsEmpty(item1) Then
        IsEqual = True
        Exit Function
    End If

    ' �����Ƃ� Null �̏ꍇ �ˈ�v
    If IsNull(item1) Then
        IsEqual = True
        Exit Function
    End If

    ' �����Ƃ� �z�� �̏ꍇ
    If IsArray(item1) Then
        If ArrayUtils.IsEqual(item1, item2) Then
            IsEqual = True
        End If

    ' �����Ƃ� Object�^ �̏ꍇ
    ElseIf IsObject(item1) Then
        If item1 Is item2 Then
            IsEqual = True
        End If

    ' �����Ƃ� ����ȊO �̏ꍇ
    Else
        If item1 = item2 Then
            IsEqual = True
        End If
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : �w�肳�ꂽ�l�𕶎���\���ɕϊ����܂��B
'
' PARAMS   : val - ������\���ɕϊ�����l
'
' RETURN   : �w�肳�ꂽ�l�̕�����\��
'
'------------------------------------------------------------------------------
Public Function ToString(ByRef val As Variant) As String
    Dim i           As Long
    Dim result      As String
    
    result = ""

    ' �z��̏ꍇ
    If IsArray(val) Then
        result = result & "("
        
        ' �v�f�����ݒ�ς݂̓��I�z��(�܂��͐ÓI�z��)�̏ꍇ
        If Not ArrayUtils.IsEmptyArray(val) Then
            result = result & ToString(val(LBound(val)))
            For i = LBound(val) + 1 To UBound(val)
                result = result & ", " & ToString(val(i))
            Next
        End If
        
        result = result & ")"
    
    ' �I�u�W�F�N�g�^�̏ꍇ
    ElseIf IsObject(val) Then
    
        On Error Resume Next
        result = val.ToString
        
        ' �I�u�W�F�N�g��ToString���\�b�h����������Ă��Ȃ��ꍇ
        If 0 < Err.number Then
            result = TypeName(val)
            Err.Clear
        End If
        
        On Error GoTo 0
    
    ' Empty�̏ꍇ
    ElseIf IsEmpty(val) Then
        result = "Empty"
    
    ' Null�̏ꍇ
    ElseIf IsNull(val) Then
        result = "Null"
    
    ' ����ȊO�̏ꍇ
    Else
        result = CStr(val)
    End If

    ToString = result

End Function
