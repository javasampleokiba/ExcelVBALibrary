VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name      As String
 
' ���O��ݒ�
Public Sub SetName(ByVal name As String)
    m_name = name
End Sub
 
' ���O���擾
Public Function GetName() As String
    GetName = m_name
End Function
 
' ���I�u�W�F�N�g�������̃I�u�W�F�N�g�Ɣ�r����
Public Function CompareTo(ByRef obj As Variant) As Integer
 
    If m_name = obj.GetName() Then
        CompareTo = 0
    ElseIf m_name < obj.GetName() Then
        CompareTo = -1
    Else
        CompareTo = 1
    End If
 
End Function

Public Function ToString() As String

    ToString = "ToString���\�b�h���Ă΂ꂽ��I"

End Function