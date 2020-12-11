Attribute VB_Name = "TestJapaneseHolidayUtils"
Option Explicit

'------------------------------------------------------------------------------
'
' MODULE : JapaneseHolidayUtils�̃e�X�g���W���[��
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

    Debug.Print "[���U]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/1/1948") = "")
    ' �K�p�J�n��
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/1/" & y) = "����")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/2/" & y) = "")
    Next

    Debug.Print "[���l�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/1948") = "")
    ' �K�p�J�n��`�j���Œ�O
    For y = 1949 To 1999
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/" & y) = "���l�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/16/" & y) = "")
    Next
    ' �j���Œ��
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/9/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/10/2000") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/7/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/8/2001") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/9/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2002") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/15/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2003") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2098") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2099") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/13/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/10/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/11/2100") = "���l�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("1/12/2100") = "")

    Debug.Print "[�����L�O�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/11/1966") = "")
    ' �K�p�J�n��
    For y = 1967 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/10/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/11/" & y) = "�����L�O�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/12/" & y) = "")
    Next

    Debug.Print "[�V�c�a����]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/1948") = "")
    ' �K�p�J�n��`���t�ύX(1���)�O
    For y = 1949 To 1988
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "�V�c�a����")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "")
    Next
    ' ���t�ύX(1���)��`���t�ύX(2���)�O
    For y = 1989 To 2018
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "�V�c�a����")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/24/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "")
    Next
    ' 2019�N�͕���
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/2019") = "")
    ' ���t�ύX(2���)��
    For y = 2020 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/23/" & y) = "�V�c�a����")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("2/24/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("12/23/" & y) = "")
    Next

    Debug.Print "[�t���̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/20/1948") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("3/21/1948") = "")
    ' �K�p�J�n��
    name = "�t���̓�"
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
    ' 2151�N�ȍ~�͌v�Z�ł��Ȃ�
    On Error Resume Next
    JapaneseHolidayUtils.GetNationalHolidayName ("3/21/2151")
    Call PrintResultIfNg(Err.number = 5)
    On Error GoTo 0

    Debug.Print "[�݂ǂ�̓�]"
    ' �K�p�J�n��`���t�ύX�O
    For y = 1989 To 2006
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "�݂ǂ�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/4/" & y) = "")
    Next
    ' ���t�ύX��
    For y = 2007 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/4/" & y) = "�݂ǂ�̓�")
    Next

    Debug.Print "[���a�̓�]"
    ' �K�p�J�n��
    For y = 2007 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/28/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/29/" & y) = "���a�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("4/30/" & y) = "")
    Next

    Debug.Print "[���@�L�O��]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/3/1948") = "")
    ' �K�p�J�n��
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/2/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/3/" & y) = "���@�L�O��")
    Next

    Debug.Print "[���ǂ��̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/5/1948") = "")
    ' �K�p�J�n��
    For y = 1949 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/5/" & y) = "���ǂ��̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("5/6/" & y) = "")
    Next

    Debug.Print "[�C�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/1995") = "")
    ' �K�p�J�n��`�j���Œ�O
    For y = 1996 To 2002
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/" & y) = "�C�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/" & y) = "")
    Next
    ' 2020�N�̂ݓ����ܗցE�p�������s�b�N���[�@�Ɋ�Â����t���قȂ�
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2020") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/23/2020") = "�C�̓�")
    ' �j���Œ��
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2003") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2004") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/17/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2005") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/16/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/17/2006") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/14/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/15/2019") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/16/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2021") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2098") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/22/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2099") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/21/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/18/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/19/2100") = "�C�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/20/2100") = "")

    Debug.Print "[�R�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/2015") = "")
    ' �K�p�J�n��
    For y = 2016 To MAX_YEAR
        If y = 2020 Then
            ' 2020�N�̂ݓ����ܗցE�p�������s�b�N���[�@�Ɋ�Â����t���قȂ�
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/9/" & y) = "")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/10/" & y) = "�R�̓�")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/" & y) = "")
        Else
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/10/" & y) = "")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/11/" & y) = "�R�̓�")
            Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("8/12/" & y) = "")
        End If
    Next

    Debug.Print "[�h�V�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/1965") = "")
    ' �K�p�J�n��`�j���Œ�O
    For y = 1966 To 2002
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/" & y) = "�h�V�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/" & y) = "")
    Next
    ' �j���Œ��
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/2003") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2004") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2004") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/18/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2005") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2005") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/17/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/18/2006") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2006") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/15/2098") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/16/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2099") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/19/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/20/2100") = "�h�V�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/21/2100") = "")

    Debug.Print "[�H���̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/22/1947") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/23/1947") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("9/24/1947") = "")
    ' �K�p�J�n��
    name = "�H���̓�"
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
    ' 2151�N�ȍ~�͌v�Z�ł��Ȃ�
    On Error Resume Next
    JapaneseHolidayUtils.GetNationalHolidayName ("9/23/2151")
    Call PrintResultIfNg(Err.number = 5)
    On Error GoTo 0

    Debug.Print "[�̈�̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/1965") = "")
    ' �K�p�J�n��`�j���Œ�O
    For y = 1966 To 1999
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/" & y) = "�̈�̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/" & y) = "")
    Next
    ' �j���Œ��
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2000") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2000") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/7/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2001") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2001") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2002") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/15/2002") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2003") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2003") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2017") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2017") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2017") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/7/2018") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2018") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2018") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2019") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2019") = "�̈�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/15/2019") = "")

    Debug.Print "[�X�|�[�c�̓�]"
    ' 2020�N�̂ݓ����ܗցE�p�������s�b�N���[�@�Ɋ�Â����t���قȂ�
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/24/2020") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("7/25/2020") = "")
    ' �K�p�J�n��
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2021") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2021") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2022") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2022") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/8/2023") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/9/2023") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2023") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2098") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/14/2098") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2099") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/13/2099") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/10/2100") = "")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/11/2100") = "�X�|�[�c�̓�")
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("10/12/2100") = "")

    Debug.Print "[�����̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/3/1947") = "")
    ' �K�p�J�n��
    For y = 1948 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/2/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/3/" & y) = "�����̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/4/" & y) = "")
    Next

    Debug.Print "[�ΘJ���ӂ̓�]"
    ' �K�p�J�n�O
    Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/23/1947") = "")
    ' �K�p�J�n��
    For y = 1948 To MAX_YEAR
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/22/" & y) = "")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/23/" & y) = "�ΘJ���ӂ̓�")
        Call PrintResultIfNg(JapaneseHolidayUtils.GetNationalHolidayName("11/24/" & y) = "")
    Next

End Sub

' TestGetNationalHolidayName�̃e�X�g�����ׂăp�X�����ꍇ�̂ݎ��s��
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

            ' ���� or �ʏ�̓y��
            If name = "" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = False)

            ' �U�֋x��
            ElseIf name = "�U�֋x��" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = False)
            
            ' �����̋x��
            ElseIf name = "�����̋x��" Then
                Call PrintResultIfNg(JapaneseHolidayUtils.IsHoliday(d) = True)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsNationalHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsSubstituteHoliday(d) = False)
                Call PrintResultIfNg(JapaneseHolidayUtils.IsCitizensHoliday(d) = True)
            
            ' �����̏j��
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
