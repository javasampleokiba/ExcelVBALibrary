Attribute VB_Name = "LangUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : ExcelVBAの共通的処理に使用されるユーティリティモジュール
'
' LINK   : このモジュールは以下のモジュールを参照しています。
'
'          ・ArrayUtils
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'
' FUNCTION : 2つの引数が同一の値であるか判定します。
'            配列の場合は配列の各要素を比較し、すべて同一の値であれば true
'            を返します。
'
' PARAMS   : item1 - 比較対象の値1
'            item2 - 比較対象の値2
'
' RETURN   : 2つの引数が同一の値である場合は true
'
'------------------------------------------------------------------------------
Public Function IsEqual(ByRef item1 As Variant, ByRef item2 As Variant) As Boolean

    IsEqual = False

    ' どちらか一方のみが Empty の場合 ⇒不一致
    If IsEmpty(item1) Xor IsEmpty(item2) Then Exit Function

    ' どちらか一方のみが Null の場合 ⇒不一致
    If IsNull(item1) Xor IsNull(item2) Then Exit Function

    ' どちらか一方のみが 配列 の場合 ⇒不一致
    If IsArray(item1) Xor IsArray(item2) Then Exit Function

    ' どちらか一方のみが Object型 の場合 ⇒不一致
    If IsObject(item1) Xor IsObject(item2) Then Exit Function

    ' 両方とも Empty の場合 ⇒一致
    If IsEmpty(item1) Then
        IsEqual = True
        Exit Function
    End If

    ' 両方とも Null の場合 ⇒一致
    If IsNull(item1) Then
        IsEqual = True
        Exit Function
    End If

    ' 両方とも 配列 の場合
    If IsArray(item1) Then
        If ArrayUtils.IsEqual(item1, item2) Then
            IsEqual = True
        End If

    ' 両方とも Object型 の場合
    ElseIf IsObject(item1) Then
        If item1 Is item2 Then
            IsEqual = True
        End If

    ' 両方とも それ以外 の場合
    Else
        If item1 = item2 Then
            IsEqual = True
        End If
    End If

End Function

'------------------------------------------------------------------------------
'
' FUNCTION : 指定された値を文字列表現に変換します。
'
' PARAMS   : val - 文字列表現に変換する値
'
' RETURN   : 指定された値の文字列表現
'
'------------------------------------------------------------------------------
Public Function ToString(ByRef val As Variant) As String
    Dim i           As Long
    Dim result      As String
    
    result = ""

    ' 配列の場合
    If IsArray(val) Then
        result = result & "("
        
        ' 要素数が設定済みの動的配列(または静的配列)の場合
        If Not ArrayUtils.IsEmptyArray(val) Then
            result = result & ToString(val(LBound(val)))
            For i = LBound(val) + 1 To UBound(val)
                result = result & ", " & ToString(val(i))
            Next
        End If
        
        result = result & ")"
    
    ' オブジェクト型の場合
    ElseIf IsObject(val) Then
    
        On Error Resume Next
        result = val.ToString
        
        ' オブジェクトにToStringメソッドが実装されていない場合
        If 0 < Err.number Then
            result = TypeName(val)
            Err.Clear
        End If
        
        On Error GoTo 0
    
    ' Emptyの場合
    ElseIf IsEmpty(val) Then
        result = "Empty"
    
    ' Nullの場合
    ElseIf IsNull(val) Then
        result = "Null"
    
    ' それ以外の場合
    Else
        result = CStr(val)
    End If

    ToString = result

End Function
