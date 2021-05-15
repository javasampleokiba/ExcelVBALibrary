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
' MODULE : 順序付けられた複数の要素を格納するコレクションクラス
'
' NOTE   : デフォルトではどんなデータ型でもリストに格納できます。
'          型安全性のため格納するデータ型を指定したい場合は、
'          このオブジェクト生成直後にSetDataTypeメソッドを実行してください。
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・ArrayUtils
'          ・LangUtils
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' メンバ変数
'------------------------------------------------------------------------------
Private dataType_       As String
Private arr_()          As Variant
Private returnItem_     As Variant

'------------------------------------------------------------------------------
' インスタンス初期化処理
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

    dataType_ = ""

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリスト内の指定位置に指定された要素を追加します。
'            位置インデックスに負数を指定するとリストの後方からの
'            位置インデックスに追加します。
'
' PARAMS   : index - 追加する位置インデックス
'            item  - 追加する要素
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Sub Add(ByVal index As Long, ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Add(arr_, index, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストからすべての要素を削除します。
'------------------------------------------------------------------------------
Public Sub Clear()

    Erase arr_

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストのコピーを生成して返します。
'
' RETURN   : コピーしたリスト
'------------------------------------------------------------------------------
Public Function Clone() As List
    Dim l   As List

    Set l = New List
    Call l.Concat(arr_)
    Set Clone = l

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストに指定された任意の数の要素を追加します。
'            要素がリストの場合は、そのリストに含まれるすべての要素を追加します。
'
' PARAMS   : items - 追加する任意の数の要素
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
' FUNCTION : 指定した要素がこのリストに含まれるか判定します。
'
' PARAMS   : item - 検索する要素
'
' RETURN   : 存在する場合は True
'------------------------------------------------------------------------------
Public Function Contains(ByRef item As Variant) As Boolean

    Contains = ArrayUtils.Contains(arr_, item)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した任意の数の要素すべてがこのリストに含まれるか判定します。
'
' PARAMS   : items - 検索する任意の数の要素
'
' RETURN   : すべての要素が存在する場合は True
'------------------------------------------------------------------------------
Public Function ContainsAll(ParamArray items() As Variant) As Boolean

    ContainsAll = ArrayUtils.ContainsAll(arr_, ArrayUtils.Flatten(2, items))

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した任意の数の要素のいずれかがこのリストに含まれるか判定します。
'
' PARAMS   : items - 検索する任意の数の要素
'
' RETURN   : いずれかの要素が存在する場合は True
'------------------------------------------------------------------------------
Public Function ContainsAny(ParamArray items() As Variant) As Boolean

    ContainsAny = ArrayUtils.ContainsAny(arr_, ArrayUtils.Flatten(2, items))

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定された要素がこのリストにいくつ存在するかカウントします。
'
' PARAMS   : item - 検索する要素
'
' RETURN   : 見つかった要素の個数
'------------------------------------------------------------------------------
Public Function Count(ByRef item As Variant) As Long

    Count = ArrayUtils.Count(arr_, item)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの先頭の要素を取得します。
'
' RETURN   : 先頭の要素
'------------------------------------------------------------------------------
Public Function First() As Variant

    If IsObjectType(ArrayUtils.First(arr_)) Then
        Set First = returnItem_
    Else
        First = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストから指定した位置インデックスの要素を取得します。
'            位置インデックスに負数を指定するとリストの後方からの
'            位置インデックスの要素を取得します。
'
' PARAMS   : index - 取得する位置インデックス
'            [def] - 位置インデックスが不正の場合に返すデフォルト値 (省略時はNull)
'
' RETURN   : 指定した位置インデックスの要素 または デフォルト値
'------------------------------------------------------------------------------
Public Function GetAt(ByVal index As Long, Optional ByRef def As Variant = Null) As Variant

    If IsObjectType(ArrayUtils.GetAt(arr_, index, def)) Then
        Set GetAt = returnItem_
    Else
        GetAt = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストに格納できるデータ型を文字列で取得します。
'
' RETURN   : このリストに格納できるデータ型
'------------------------------------------------------------------------------
Public Function GetDataType() As String

    GetDataType = dataType_

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素がこのリスト内で最初に見つかった位置インデックスを
'            返します。
'
' PARAMS   : item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'------------------------------------------------------------------------------
Public Function IndexOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long

    IndexOf = ArrayUtils.IndexOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素がこのリスト内で見つかったすべての位置インデックスを
'            返します。
'
' PARAMS   : item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'------------------------------------------------------------------------------
Public Function IndicesOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long()

    IndicesOf = ArrayUtils.IndicesOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストに要素がない場合にTrueを返します。
'
' RETURN   : このリストに要素がない場合はTrue
'------------------------------------------------------------------------------
Public Function IsEmptyList() As Boolean

    IsEmptyList = ArrayUtils.IsEmptyArray(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストと指定されたリストが等しいか判定します。
'
' PARAMS   : other - 比較対象のリスト
'
' RETURN   : 指定されたリストと等しい場合は True
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef other As List) As Boolean

    IsEqual = ArrayUtils.IsEqual(arr_, other.ToArray())

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの各要素の文字列表現を指定された区切り文字で連結し、
'            返します。
'
' PARAMS   : [sep] - 区切り文字 (省略された場合は空文字)
'
' RETURN   : 連結された文字列
'------------------------------------------------------------------------------
Public Function Join(Optional ByVal sep As String = "") As String

    Join = ArrayUtils.Join(arr_, sep)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの最後の要素を取得します。
'
' RETURN   : 最後の要素
'------------------------------------------------------------------------------
Public Function Last() As Variant

    If IsObjectType(ArrayUtils.Last(arr_)) Then
        Set Last = returnItem_
    Else
        Last = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : 指定した要素がこのリスト内で最後に見つかった位置インデックスを
'            返します。
'
' PARAMS   : item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックス (存在しない場合は -1)
'------------------------------------------------------------------------------
Public Function LastIndexOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long

    LastIndexOf = ArrayUtils.LastIndexOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : 最後から検索して、指定した要素がこのリスト内で見つかった
'            すべての位置インデックスを返します。
'
' PARAMS   : item    - 検索する要素
'            [start] - 検索開始位置インデックス (範囲外でも正しく動作します)
'
' RETURN   : 指定した要素が見つかった位置インデックスの配列
'            (存在しない場合は空の配列)
'------------------------------------------------------------------------------
Public Function LastIndicesOf(ByRef item As Variant, Optional ByVal start As Long = -1) As Long()

    LastIndicesOf = ArrayUtils.LastIndicesOf(arr_, item, start)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリスト内にある要素の数を返します。
'
' RETURN   : 要素の数
'------------------------------------------------------------------------------
Public Function Length() As Long

    Length = ArrayUtils.Length(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの中で最大の要素を取得します。
'            すべての要素が不等号演算子による比較が可能であることが前提条件です。
'
' RETURN   : 最大の要素
'------------------------------------------------------------------------------
Public Function Max() As Variant

    If IsObjectType(ArrayUtils.Max(arr_)) Then
        Set Max = returnItem_
    Else
        Max = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの中で最小の要素を取得します。
'            すべての要素が不等号演算子による比較が可能であることが前提条件です。
'
' RETURN   : 最小の要素
'------------------------------------------------------------------------------
Public Function Min() As Variant

    If IsObjectType(ArrayUtils.Min(arr_)) Then
        Set Min = returnItem_
    Else
        Min = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの最後の要素を削除して返します。
'
' RETURN   : 削除した要素
'------------------------------------------------------------------------------
Public Function Pop() As Variant

    If IsObjectType(ArrayUtils.Pop(arr_)) Then
        Set Pop = returnItem_
    Else
        Pop = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの最後に要素を追加します。
'
' PARAMS   : item - 追加する要素
'------------------------------------------------------------------------------
Public Sub Push(ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Push(arr_, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストから指定した要素を削除します。
'
' PARAMS   : item   - 削除する要素
'            [size] - 削除する最大数 (省略すると、すべて削除)
'------------------------------------------------------------------------------
Public Sub Remove(ByRef item As Variant, Optional ByVal size As Long = 0)

    Call ArrayUtils.Remove(arr_, item, size)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストから指定した任意の数の要素すべてを削除します。
'
' PARAMS   : items  - 削除する任意の数の要素
'------------------------------------------------------------------------------
Public Sub RemoveAll(ParamArray items() As Variant)

    Call ArrayUtils.RemoveAll(arr_, items)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリスト内の指定した位置の要素を削除します。
'            位置インデックスに負数を指定するとリストの後方からの
'            位置インデックスの要素を削除します。
'
' PARAMS   : index - 削除する位置インデックス
'
' RETURN   : 削除された要素
'
' ERROR    : 位置インデックスが配列の範囲外の場合
'------------------------------------------------------------------------------
Public Function RemoveAt(ByVal index As Long) As Variant

    If IsObjectType(ArrayUtils.RemoveAt(arr_, index)) Then
        Set RemoveAt = returnItem_
    Else
        RemoveAt = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの先頭から指定した要素を検索し、指定した別の要素に
'            置き換えます。
'
' PARAMS   : item        - 置換対象要素
'            replacement - 置換する要素
'            [size]      - 置換する最大数 (省略すると、すべて置換)
'
' RETURN   : 置換した要素数
'------------------------------------------------------------------------------
Public Function Replace(ByRef item As Variant, ByRef replacement As Variant, _
                        Optional ByVal size As Long = 0) As Long

    Call CheckDataType(replacement)

    Replace = ArrayUtils.Replace(arr_, item, replacement, size)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの末尾から指定した要素を検索し、指定した別の要素に置き換えます。
'
' PARAMS   : item        - 置換対象要素
'            replacement - 置換する要素
'            [size]      - 置換する最大数 (省略すると、すべて置換)
'
' RETURN   : 置換した要素数
'------------------------------------------------------------------------------
Public Function ReplaceLast(ByRef item As Variant, ByRef replacement As Variant, _
                            Optional ByVal size As Long = 0) As Long

    Call CheckDataType(replacement)

    ReplaceLast = ArrayUtils.ReplaceLast(arr_, item, replacement, size)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素の順番を反転します。
'------------------------------------------------------------------------------
Public Sub Reverse()

    Call ArrayUtils.Reverse(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素を指定された距離だけ回転させます。
'            距離に正の値を指定すると要素はリストの後ろ方向に移動し、
'            負の値を指定すると前方向に移動します。
'
' PARAMS   : distance - 要素の移動距離
'------------------------------------------------------------------------------
Public Sub Rotate(ByVal distance As Long)

    Call ArrayUtils.Rotate(arr_, distance)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素を指定個数ランダム格納したリストを返します。
'
' PARAMS   : [size] - 取得する個数 (省略すると1つ。最大は配列サイズ)
'            [uniq] - 同じ位置の要素を複数取得不可とするかのフラグ
'
' RETURN   : 選ばれた要素から成る新しいリスト
'------------------------------------------------------------------------------
Public Function Sample(Optional ByVal size As Long = 1, _
                       Optional ByVal uniq As Boolean = True) As List
    Dim arr()   As Variant

    arr = ArrayUtils.Sample(arr_, size, uniq)
    Set Sample = NewList(arr)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの指定した位置インデックスに要素を設定します。
'            位置インデックスに負数を指定するとリストの後方からの
'            位置インデックスに追加します。
'
' PARAMS   : index - 設定する位置インデックス
'            item  - 設定する要素
'
' RETURN   : 設定前に指定した位置インデックスにあった要素
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
' FUNCTION : このリストに格納できるデータ型を文字列で指定します。
'
' PARAMS   : dataType - データ型
'
' RETURN   : リストが空の場合は常にTrueを返します。
'            すでにデータが格納されており、指定されたデータ型と
'            一致しない場合はFalseを返します。
'------------------------------------------------------------------------------
Public Function SetDataType(ByVal dataType As String) As Boolean
    Dim item        As Variant

    ' 配列が空ではない場合は設定する前に先にチェックを行う
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
' FUNCTION : このリストの先頭の要素を削除して返します。
'
' RETURN   : 削除した要素
'------------------------------------------------------------------------------
Public Function Shift() As Variant

    If IsObjectType(ArrayUtils.Shift(arr_)) Then
        Set Shift = returnItem_
    Else
        Shift = returnItem_
    End If

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素の順番をランダムに入れ替えます。
'------------------------------------------------------------------------------
Public Sub Shuffle()

    Call ArrayUtils.Shuffle(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストから引数で指定した範囲の要素を取り出し、
'            新しいリストを生成して返します。
'
' PARAMS   : [startIdx] - 取得する開始位置インデックス (省略すると先頭)
'            [endIdx]   - 取得する終了位置インデックス。ただし、endIdxの位置は
'                         含まれない。 (省略すると末尾まで取得)
'
' RETURN   : 取り出した要素から成る新しいリスト
'------------------------------------------------------------------------------
Public Function Slice(Optional ByVal startIdx As Long = 0, _
                       Optional ByVal endIdx As Long = 0) As List

    Set Slice = NewList(ArrayUtils.Slice(arr_, startIdx, endIdx))

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素を昇順にソートします。
'            すべての要素が不等号演算子による比較が可能であることが前提条件です。
'------------------------------------------------------------------------------
Public Sub Sort()

    Call ArrayUtils.Sort(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストから要素を取り除きつつ、新しい要素を追加します。
'            追加する要素を指定しない場合は単に要素を削除します。
'            削除数を0にすると単に要素を挿入します。
'
' PARAMS   : index     - リストを変化させ始めるインデックス
'                        (負数を指定するとリストの後方からの位置インデックス)
'            [size]    - 削除する最大要素数 (省略すると開始位置から後ろの全要素を削除)
'            [addList] - 追加する要素を格納したリスト
'
' ERROR    : 位置インデックスがリストの範囲外の場合
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
' FUNCTION : このリスト内の2つの要素の位置を入れ替えます。
'
' PARAMS   : index1 - 変更する位置インデックス1
'            index2 - 変更する位置インデックス2
'------------------------------------------------------------------------------
Public Sub Swap(ByVal index1 As Long, ByVal index2 As Long)

    Call ArrayUtils.Swap(arr_, index1, index2)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストの要素で構成される新しい配列を生成して返します。
'
' RETURN   : このリストの要素で構成される新しい配列
'------------------------------------------------------------------------------
Public Function ToArray() As Variant()

    ToArray = arr_

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストの文字列表現を返します。
'
' RETURN   : このリストの文字列表現
'------------------------------------------------------------------------------
Public Function ToString() As String

    ToString = LangUtils.ToString(arr_)

End Function

'------------------------------------------------------------------------------
' FUNCTION : このリストから重複する要素をすべて削除します。
'------------------------------------------------------------------------------
Public Sub Unique()

    Call ArrayUtils.Unique(arr_)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストの先頭に要素を追加します。
'
' PARAMS   : item - 追加する要素
'------------------------------------------------------------------------------
Public Sub Unshift(ByRef item As Variant)

    Call CheckDataType(item)

    Call ArrayUtils.Unshift(arr_, item)

End Sub

'------------------------------------------------------------------------------
' FUNCTION : このリストから引数で指定した位置の要素を取り出し、
'            新しいリストを生成して返します。
'
' PARAMS   : indices - 取得する位置インデックス、または位置インデックスの配列
'
' RETURN   : 取り出した要素から成る新しいリスト
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

