Attribute VB_Name = "TestJapaneseHolidayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : JapaneseHolidayUtilsのテストモジュール
'
'------------------------------------------------------------------------------

Public Sub TestAll()

    Debug.Print "=== TestJapaneseHolidayUtils ==="

    Call TestGetNationalHolidayName
    Call TestOthers

End Sub

Private Sub TestGetNationalHolidayName()
    Dim MAX_YEAR    As Integer: MAX_YEAR = 3000
    Dim y           As Integer
    Dim name        As String

    Debug.Print "--- TestGetNationalHolidayName ---"

    Debug.Print "[元旦]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/1/1948") = "")
    ' 適用開始後
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/1/" & y) = "元日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/2/" & y) = "")
    Next

    Debug.Print "[成人の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/1948") = "")
    ' 適用開始後～曜日固定前
    For y = 1949 To 1999
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/" & y) = "成人の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/16/" & y) = "")
    Next
    ' 曜日固定後
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/9/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/10/2000") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/7/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/8/2001") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/9/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2002") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2003") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2098") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2099") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/10/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2100") = "成人の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2100") = "")

    Debug.Print "[建国記念の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/11/1966") = "")
    ' 適用開始後
    For y = 1967 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/10/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/11/" & y) = "建国記念の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/12/" & y) = "")
    Next

    Debug.Print "[天皇誕生日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/1948") = "")
    ' 適用開始後～日付変更(1回目)前
    For y = 1949 To 1988
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "天皇誕生日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "")
    Next
    ' 日付変更(1回目)後～日付変更(2回目)前
    For y = 1989 To 2018
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "天皇誕生日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/24/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "")
    Next
    ' 2019年は平日
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/2019") = "")
    ' 日付変更(2回目)後
    For y = 2020 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "天皇誕生日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/24/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "")
    Next

    Debug.Print "[春分の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1948") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1948") = "")
    ' 適用開始後
    name = "春分の日"
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1949") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1950") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1951") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1952") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1953") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1954") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1955") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1956") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1957") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1958") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1959") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1960") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1961") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1962") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1963") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1964") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1965") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1966") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1967") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1968") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1969") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1970") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1971") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1972") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1973") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1974") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1975") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1976") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1977") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1978") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1979") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1980") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1981") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1982") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1983") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1984") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1985") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1986") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1987") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1988") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1989") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1990") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1991") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1992") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1993") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1994") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1995") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1996") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1997") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1998") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1999") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2000") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2001") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2002") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2003") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2004") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2005") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2006") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2007") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2008") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2009") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2010") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2011") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2012") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2013") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2014") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2015") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2016") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2017") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2018") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2019") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2020") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2021") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2022") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2023") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2024") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2025") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2026") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2027") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2028") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2029") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2030") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2031") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2032") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2033") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2034") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2035") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2036") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2037") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2038") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2039") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2040") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2041") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2042") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2043") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2044") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2045") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2046") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2047") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2048") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2049") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2050") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2141") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2142") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2143") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2144") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2145") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2146") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2147") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2148") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/2149") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/2150") = name)
    ' 2151年以降は計算できない
    On Error Resume Next
    JapaneseHolidayUtils.GetNationalHolidayName ("3/21/2151")
    Call PrintResultIfNg(Err.Number = 5)
    On Error GoTo 0

    Debug.Print "[みどりの日]"
    ' 適用開始後～日付変更前
    For y = 1989 To 2006
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "みどりの日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/4/" & y) = "")
    Next
    ' 日付変更後
    For y = 2007 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/4/" & y) = "みどりの日")
    Next

    Debug.Print "[昭和の日]"
    ' 適用開始後
    For y = 2007 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "昭和の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
    Next

    Debug.Print "[憲法記念日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/3/1948") = "")
    ' 適用開始後
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/2/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/3/" & y) = "憲法記念日")
    Next

    Debug.Print "[こどもの日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/5/1948") = "")
    ' 適用開始後
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/5/" & y) = "こどもの日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/6/" & y) = "")
    Next

    Debug.Print "[海の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/1995") = "")
    ' 適用開始後～曜日固定前
    For y = 1996 To 2002
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/" & y) = "海の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/" & y) = "")
    Next
    ' 東京五輪・パラリンピック特措法に基づき2020年は日付が異なる
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2020") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/23/2020") = "海の日")
    ' 改正 東京五輪・パラリンピック特措法に基づき2021年は日付が異なる
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2021") = "海の日")
    ' 曜日固定後
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2003") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2004") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/17/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2005") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/16/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/17/2006") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/14/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/15/2019") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/16/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/17/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2022") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2098") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2099") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2100") = "海の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2100") = "")

    Debug.Print "[山の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/2015") = "")
    ' 適用開始後
    For y = 2016 To MAX_YEAR
        If y = 2020 Then
            ' 東京五輪・パラリンピック特措法に基づき2020年は日付が異なる
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/9/" & y) = "")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/10/" & y) = "山の日")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/" & y) = "")
        ElseIf y = 2021 Then
            ' 改正 東京五輪・パラリンピック特措法に基づき2021年は日付が異なる
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/7/" & y) = "")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/8/" & y) = "山の日")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/9/" & y) = "")
        Else
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/10/" & y) = "")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/" & y) = "山の日")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/12/" & y) = "")
        End If
    Next

    Debug.Print "[敬老の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/1965") = "")
    ' 適用開始後～曜日固定前
    For y = 1966 To 2002
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/" & y) = "敬老の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/" & y) = "")
    Next
    ' 曜日固定後
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/2003") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2004") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/18/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2005") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/17/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/18/2006") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/2098") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2099") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2100") = "敬老の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2100") = "")

    Debug.Print "[秋分の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/1947") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1947") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1947") = "")
    ' 適用開始後
    name = "秋分の日"
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1948") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1949") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1950") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1951") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1952") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1953") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1954") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1955") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1956") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1957") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1958") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1959") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1960") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1961") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1962") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1963") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1964") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1965") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1966") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1967") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1968") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1969") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1970") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1971") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1972") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1973") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1974") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1975") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1976") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1977") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1978") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1979") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1980") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1981") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1982") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1983") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1984") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1985") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1986") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1987") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1988") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1989") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1990") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1991") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1992") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1993") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1994") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1995") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1996") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1997") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1998") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1999") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2000") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2001") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2002") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2003") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2004") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2005") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2006") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2007") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2008") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2009") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2010") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2011") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2012") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2013") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2014") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2015") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2016") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2017") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2018") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2019") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2020") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2021") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2022") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2023") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2024") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2025") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2026") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2027") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2028") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2029") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2030") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2031") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2032") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2033") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2034") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2035") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2036") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2037") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2038") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2039") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2040") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2041") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2042") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2043") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2044") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2045") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2046") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2047") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2048") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2049") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2050") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2141") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2142") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2143") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2144") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2145") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2146") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2147") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2148") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2149") = name)
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/2150") = name)
    ' 2151年以降は計算できない
    On Error Resume Next
    JapaneseHolidayUtils.GetNationalHolidayName ("9/23/2151")
    Call PrintResultIfNg(Err.Number = 5)
    On Error GoTo 0

    Debug.Print "[体育の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/1965") = "")
    ' 適用開始後～曜日固定前
    For y = 1966 To 1999
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/" & y) = "体育の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/" & y) = "")
    Next
    ' 曜日固定後
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2000") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/7/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2001") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2002") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/15/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2003") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2017") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2017") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2017") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/7/2018") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2018") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2018") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2019") = "体育の日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/15/2019") = "")

    Debug.Print "[スポーツの日]"
    ' 東京五輪・パラリンピック特措法に基づき2020年は日付が異なる
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/24/2020") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/25/2020") = "")
    ' 改正 東京五輪・パラリンピック特措法に基づき2021年は日付が異なる
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/23/2021") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/24/2021") = "")
    ' 適用開始後
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2021") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2022") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2023") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2023") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2023") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2098") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2099") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2100") = "スポーツの日")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2100") = "")

    Debug.Print "[文化の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/3/1947") = "")
    ' 適用開始後
    For y = 1948 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/2/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/3/" & y) = "文化の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/4/" & y) = "")
    Next

    Debug.Print "[勤労感謝の日]"
    ' 適用開始前
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/23/1947") = "")
    ' 適用開始後
    For y = 1948 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/23/" & y) = "勤労感謝の日")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/24/" & y) = "")
    Next

End Sub

' TestGetNationalHolidayNameのテストをすべてパスした場合のみ実行可
Private Sub TestOthers()
    Dim d       As Date
    Dim y       As Integer
    Dim dy      As Integer
    Dim name    As String

    Debug.Print "--- TestOthers ---"

    For y = 1947 To 2050
        d = "1/1/" & y
        For dy = 0 To 365
            name = JapaneseHolidayUtils.GetHolidayName(d)

            ' 平日 or 通常の土日
            If name = "" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = False)

            ' 振替休日
            ElseIf name = "振替休日" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = False)
            
            ' 国民の休日
            ElseIf name = "国民の休日" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = True)
            
            ' 国民の祝日
            Else
                Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName(d) = name)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = False)

            End If

            d = DateAdd("d", 1, d)
            If year(d) <> y Then
                Exit For
            End If
        Next
    Next

End Sub
