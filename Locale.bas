Attribute VB_Name = "Locale"
Option Explicit
Private Const LOCALE_USER_DEFAULT& = &H400
Private Const LOCALE_SDECIMAL& = &HE
Private Const LOCALE_STHOUSAND& = &HF
Private Declare Function GetLocaleInfo& Lib "KERNEL32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long)

Public Function ThousandSeparator() As String
    Dim r As Long, s As String
    s = String(10, "a")
    r = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, s, 10)
    ThousandSeparator = Left$(s, r)
End Function

Public Function DecimalSeparator() As String
    Dim r As Long, s As String
    s = String(10, "a")
    r = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, s, 10)
    DecimalSeparator = Left$(s, r)
End Function
