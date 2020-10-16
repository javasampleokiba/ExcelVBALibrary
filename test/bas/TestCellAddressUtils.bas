Attribute VB_Name = "TestCellAddressUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : CellAddressUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll()

    Debug.Print "=== TestCellAddressUtils ==="
    
    Call TestToColumnIndex
    Call TestToColumnName

End Sub

Private Sub TestToColumnIndex()

    Debug.Print "--- TestToColumnIndex ---"

    Call PrintResult(CellAddressUtils.ToColumnIndex("A") = 1, 1)
    Call PrintResult(CellAddressUtils.ToColumnIndex("z") = 26, 2)
    Call PrintResult(CellAddressUtils.ToColumnIndex("Aa") = 27, 3)
    Call PrintResult(CellAddressUtils.ToColumnIndex("ZZ") = 702, 4)
    Call PrintResult(CellAddressUtils.ToColumnIndex("aaa") = 703, 5)
    Call PrintResult(CellAddressUtils.ToColumnIndex("XFD") = 16384, 6)
    Call PrintResult(CellAddressUtils.ToColumnIndex("zZZz") = 475254, 7)
    Call PrintResult(CellAddressUtils.ToColumnIndex("AAAAA") = 475255, 8)
    Call PrintResult(CellAddressUtils.ToColumnIndex("A1") = 0, 9)
    Call PrintResult(CellAddressUtils.ToColumnIndex("A-Z") = 0, 10)

End Sub

Private Sub TestToColumnName()

    Debug.Print "--- TestToColumnName ---"

    Call PrintResult(CellAddressUtils.ToColumnName(1) = "A", 1)
    Call PrintResult(CellAddressUtils.ToColumnName(26) = "Z", 2)
    Call PrintResult(CellAddressUtils.ToColumnName(27) = "AA", 3)
    Call PrintResult(CellAddressUtils.ToColumnName(702) = "ZZ", 4)
    Call PrintResult(CellAddressUtils.ToColumnName(703) = "AAA", 5)
    Call PrintResult(CellAddressUtils.ToColumnName(16384) = "XFD", 6)
    Call PrintResult(CellAddressUtils.ToColumnName(475254) = "ZZZZ", 7)
    Call PrintResult(CellAddressUtils.ToColumnName(475255) = "AAAAA", 8)
    Call PrintResult(CellAddressUtils.ToColumnName(0) = "", 9)
    Call PrintResult(CellAddressUtils.ToColumnName(-1) = "", 10)

End Sub
