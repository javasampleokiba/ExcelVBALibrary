VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' MODULE : 各種反復処理可能なデータのイテレーションを行うクラス
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・ArrayUtils
'          ・LangUtils
'          ・List
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' メンバ変数
'------------------------------------------------------------------------------
Private Enum IteratorType
    TYPE_ARRAY = 1
    TYPE_LIST = 2
End Enum

Private itType_         As Integer
Private curIdx_         As Long
Private count_          As Long
Private arr_            As Variant
Private list_           As List
Private mode_           As String
Private objectFlg_      As Boolean
Private returnItem_     As Variant

'------------------------------------------------------------------------------
' FUNCTION : このイテレータを指定された配列のイテレータに変更、初期化します。
'            modeでイテレートする方法を指定できます。
'            Reverse：逆順にイレテートする
'
' PARAMS   : arr    - イテレートする配列
'            [mode] - イテレートする方法
'------------------------------------------------------------------------------
Public Sub InitArrayIterator(ByRef arr As Variant, Optional ByVal mode As String)

    itType_ = IteratorType.TYPE_ARRAY
    arr_ = arr
    mode_ = LCase(mode)
    count_ = 0

    If ArrayUtils.IsEmptyArray(arr_) Then
        curIdx_ = -1
    ElseIf mode_ = "reverse" Then
        curIdx_ = UBound(arr_) + 1
    Else
        curIdx_ = LBound(arr_) - 1
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このイテレータを指定されたListのイテレータに変更、初期化します。
'            modeでイテレートする方法を指定できます。
'            Reverse：逆順にイレテートする
'
' PARAMS   : arr    - イテレートする配列
'            [mode] - イテレートする方法
'------------------------------------------------------------------------------
Public Sub InitListIterator(ByRef lst As List, Optional ByVal mode As String)

    itType_ = IteratorType.TYPE_LIST
    Set list_ = lst
    mode_ = LCase(mode)
    count_ = 0

    If mode_ = "reverse" Then
        curIdx_ = list_.Length()
    Else
        curIdx_ = -1
    End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 次の要素があるか判定します。
'
' RETURN   : 次の要素がある場合は TRUE
'------------------------------------------------------------------------------
Public Function HasNext() As Boolean
    Dim nextIdx     As Long

    nextIdx = NextIndex()

    Select Case itType_
    Case IteratorType.TYPE_ARRAY
        HasNext = HasNextArray(nextIdx)
    Case IteratorType.TYPE_LIST
        HasNext = HasNextList(nextIdx)
    End Select

End Function

'------------------------------------------------------------------------------
' FUNCTION : 次の要素を取得します。
'
' RETURN   : 次の要素
'------------------------------------------------------------------------------
Public Function GetNext() As Variant

    Call SetNextIndex

    Select Case itType_
    Case IteratorType.TYPE_ARRAY
        Call NextArray
    Case IteratorType.TYPE_LIST
        Call NextList
    End Select

    If objectFlg_ Then
        Set GetNext = returnItem_
    Else
        GetNext = returnItem_
    End If

    count_ = count_ + 1

End Function

'------------------------------------------------------------------------------
' FUNCTION : このイテレータによって最後に返された要素を削除します。
'
' ERROR    : この操作に対応していない場合
'------------------------------------------------------------------------------
Public Sub Remove()

    Select Case itType_
    Case IteratorType.TYPE_ARRAY
        ' 配列は削除に非対応
        Call Err.Raise(17)
    Case IteratorType.TYPE_LIST
        Call RemoveList
    End Select

    Call SetPrevIndex

End Sub

'------------------------------------------------------------------------------
' FUNCTION : 現在のインデックスを取得します。
'
' RETURN   : 現在のインデックス
'------------------------------------------------------------------------------
Public Function GetCurrentIndex() As Long

    GetCurrentIndex = curIdx_

End Function

'------------------------------------------------------------------------------
' FUNCTION : 反復回数(GetNextメソッドが何回実行されたか)を取得します。
'
' RETURN   : 反復回数
'------------------------------------------------------------------------------
Public Function GetCount() As Long

    GetCount = count_

End Function

'------------------------------------------------------------------------------
' FUNCTION : 直前のGetNextメソッドで取得した要素が最初の要素であるか判定します。
'            (この結果は、GetCount() = 1の結果と常に同じです)
'
' RETURN   : 最初の要素の場合は True
'------------------------------------------------------------------------------
Public Function IsFirst() As Long

    If GetCount() = 1 Then
        IsFirst = True
    Else
        IsFirst = False
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 直前のGetNextメソッドで取得した要素が最後の要素であるか判定します。
'            (この結果は、Not HasNext()の結果と常に同じです)
'
' RETURN   : 最初の要素の場合は True
'------------------------------------------------------------------------------
Public Function IsLast() As Long

    If Not HasNext() Then
        IsLast = True
    Else
        IsLast = False
    End If

End Function

Private Function NextIndex() As Long

    If mode_ = "reverse" Then
        NextIndex = curIdx_ - 1
    Else
        NextIndex = curIdx_ + 1
    End If

End Function

Private Sub SetNextIndex()

    If mode_ = "reverse" Then
        curIdx_ = curIdx_ - 1
    Else
        curIdx_ = curIdx_ + 1
    End If

End Sub

Private Sub SetPrevIndex()

    If mode_ <> "reverse" Then
        curIdx_ = curIdx_ - 1
    End If

End Sub

Private Sub SetNext(ByRef item As Variant)

    If IsObject(item) Then
        objectFlg_ = True
        Set returnItem_ = item
    Else
        objectFlg_ = False
        returnItem_ = item
    End If

End Sub

'------------------------------------------------------------------------------
' Methods for Array
'------------------------------------------------------------------------------
Private Function HasNextArray(ByVal nextIdx As Long) As Boolean
    Dim idx     As Long

    If ArrayUtils.IsEmptyArray(arr_) Then
        HasNextArray = False
        Exit Function
    ElseIf nextIdx < LBound(arr_) Or UBound(arr_) < nextIdx Then
        HasNextArray = False
        Exit Function
    End If

    HasNextArray = True

End Function

Private Sub NextArray()

    Call SetNext(arr_(curIdx_))

End Sub

Private Sub RemoveArray()

    Call Err.Raise(17)

End Sub

'------------------------------------------------------------------------------
' Methods for List
'------------------------------------------------------------------------------
Private Function HasNextList(ByVal nextIdx As Long) As Boolean
    Dim idx     As Long

    If list_.IsEmptyList() Then
        HasNextList = False
        Exit Function
    ElseIf nextIdx < 0 Or list_.Length() <= nextIdx Then
        HasNextList = False
        Exit Function
    End If

    HasNextList = True

End Function

Private Sub NextList()

    Call SetNext(list_.GetAt(curIdx_))

End Sub

Private Sub RemoveList()

    Call list_.RemoveAt(curIdx_)

End Sub
