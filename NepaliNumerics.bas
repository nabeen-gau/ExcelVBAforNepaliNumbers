'Declare Global variables for storing numbers from 0 to 99 and hundred to kharab
Dim NepaliNumbers() As Variant
Dim Specials() As Variant

Function GetNepali(ByVal Number As LongLong) As String
    Dim FirstDigit As LongLong
    Dim RemainingDigits As LongLong
    Dim result As String
       If Number >= 0 And Number <= 99 Then ' zero to hundred
            result = NepaliNumbers(CLng(Number))
        
        ElseIf Number >= 100 And Number <= 999 Then ' hundred to thousand
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(0)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(0) & NepaliNumbers(CLng(RemainingDigits))
            End If
            
        ElseIf Number >= 1000 And Number <= 9999 Then ' thousand to ten thousand
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(1)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(1) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 10000 And Number <= 99999 Then ' ten thousand to one lakh
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 2))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 2))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(1)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(1) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 100000 And Number <= 999999 Then ' one lakh to 10 lakh
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(2)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(2) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 1000000 And Number <= 9999999 Then '10 lakh to 99 lakh
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 2))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 2))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(2)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(2) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 10000000 And Number <= 99999999 Then '1 crore to 10 crore
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(3)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(3) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 100000000 And Number <= 999999999 Then '10 crore to 1 arab
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 2))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 2))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(3)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(3) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 1000000000 And Number <= 9999999999# Then ' 1 arab to 10 arab
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(4)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(4) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 10000000000# And Number <= 99999999999# Then '10 arab to 1 kharab
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 2))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 2))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(4)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(4) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 100000000000# And Number <= 999999999999# Then ' 1 arab to 10 arab
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 1))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 1))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(5)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(5) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 1000000000000# And Number <= 9999999999999# Then '10 arab to 1 kharab
            FirstDigit = Number \ (10 ^ (Len(CStr(Number)) - 2))
            RemainingDigits = Number Mod (10 ^ (Len(CStr(Number)) - 2))
            If RemainingDigits = 0 Then
                result = GetNepali(CLng(FirstDigit)) & Specials(5)
            Else
                result = NepaliNumbers(CLng(FirstDigit)) & Specials(5) & GetNepali(RemainingDigits)
            End If
        ElseIf Number >= 10000000000000# Then ' greater than 1 kharab
            FirstDigit = Number \ (10 ^ (11))
            RemainingDigits = Number Mod (10 ^ (11))
            If RemainingDigits = 0 Then
                result = GetNepali(FirstDigit) & Specials(5)
            Else
                result = GetNepali(FirstDigit) & Specials(5) & GetNepali(RemainingDigits)
            End If
        Else
            result = "Not Implemented"
        End If
        GetNepali = result
End Function

Function ConvertToNepaliText(ByVal Number As LongLong) As String
    
    NepaliNumbers = Array(ChrW(2358) & ChrW(2370) & ChrW(2344) & ChrW(2381) & ChrW(2351), ChrW(2319) & ChrW(2325), ChrW(2342) & ChrW(2369) & ChrW(2312), ChrW(2340) & ChrW(2367) & ChrW(2344), ChrW(2330) & ChrW(2366) & ChrW(2352), ChrW(2346) & ChrW(2366) & ChrW(2305) & ChrW(2330), ChrW(2331), ChrW(2360) & ChrW(2366) & ChrW(2340), ChrW(2310) & ChrW(2336), ChrW(2344) & ChrW(2380) & ChrW(2305), ChrW(2342) & ChrW(2358), ChrW(2319) & ChrW(2328) & ChrW(2366) & ChrW(2352), ChrW(2348) & ChrW(2366) & ChrW(2361) & ChrW(2381) & ChrW(2352), ChrW(2340) & ChrW(2375) & ChrW(2361) & ChrW(2381) & ChrW(2352), ChrW(2330) & ChrW(2380) & ChrW(2343), ChrW(2346) & ChrW(2344) & ChrW(2381) & ChrW(2342) & ChrW(2381) & ChrW(2352), _
						ChrW(2360) & ChrW(2379) & ChrW(2361) & ChrW(2381) & ChrW(2352), ChrW(2360) & ChrW(2340) & ChrW(2381) & ChrW(2352), ChrW(2309) & ChrW(2336) & ChrW(2366) & ChrW(2352), ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344) & ChrW(2366) & ChrW(2311) & ChrW(2360), ChrW(2348) & ChrW(2367) & ChrW(2360), ChrW(2319) & ChrW(2325) & ChrW(2381) & ChrW(2325) & ChrW(2366) & ChrW(2311) & ChrW(2360), ChrW(2348) & ChrW(2366) & ChrW(2311) & ChrW(2360), ChrW(2340) & ChrW(2375) & ChrW(2311) & ChrW(2360), ChrW(2330) & ChrW(2380) & ChrW(2348) & ChrW(2367) & ChrW(2360), ChrW(2346) & ChrW(2330) & ChrW(2381) & ChrW(2330) & ChrW(2367) & ChrW(2360), ChrW(2331) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2367) & ChrW(2360), _
						ChrW(2360) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2366) & ChrW(2311) & ChrW(2360), ChrW(2309) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2366) & ChrW(2311) & ChrW(2360), ChrW(2313) & ChrW(2344) & ChrW(2344) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2319) & ChrW(2325) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2348) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2340) & ChrW(2375) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2330) & ChrW(2380) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2346) & ChrW(2376) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2331) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), _
						ChrW(2360) & ChrW(2337) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2309) & ChrW(2337) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2360), ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2319) & ChrW(2325) & ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2348) & ChrW(2351) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2330) & ChrW(2380) & ChrW(2357) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2346) & ChrW(2376) & ChrW(2340) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), _
						ChrW(2331) & ChrW(2351) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2360) & ChrW(2337) & ChrW(2381) & ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2309) & ChrW(2337) & ChrW(2381) & ChrW(2330) & ChrW(2366) & ChrW(2354) & ChrW(2367) & ChrW(2360), ChrW(2313) & ChrW(2344) & ChrW(2344) & ChrW(2381) & ChrW(2346) & ChrW(2330) & ChrW(2366) & ChrW(2360), ChrW(2346) & ChrW(2330) & ChrW(2366) & ChrW(2360), ChrW(2319) & ChrW(2325) & ChrW(2366) & ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2348) & ChrW(2366) & ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2330) & ChrW(2380) & ChrW(2357) & ChrW(2344) & ChrW(2381) & ChrW(2344), _
						ChrW(2346) & ChrW(2330) & ChrW(2381) & ChrW(2346) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2331) & ChrW(2346) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2360) & ChrW(2344) & ChrW(2381) & ChrW(2340) & ChrW(2366) & ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2309) & ChrW(2344) & ChrW(2381) & ChrW(2336) & ChrW(2366) & ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2360) & ChrW(2366) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2360) & ChrW(2366) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2319) & ChrW(2325) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2348) & ChrW(2376) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), _
						ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2330) & ChrW(2380) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2346) & ChrW(2376) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2331) & ChrW(2376) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2360) & ChrW(2337) & ChrW(2381) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2309) & ChrW(2337) & ChrW(2360) & ChrW(2336) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2360) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352) & ChrW(2368), ChrW(2360) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352) & ChrW(2368), ChrW(2319) & ChrW(2325) & ChrW(2340) & ChrW(2352), _
						ChrW(2348) & ChrW(2361) & ChrW(2340) & ChrW(2352), ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2340) & ChrW(2352), ChrW(2330) & ChrW(2380) & ChrW(2352) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352), ChrW(2346) & ChrW(2330) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352), ChrW(2331) & ChrW(2351) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352), ChrW(2360) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2340) & ChrW(2352), ChrW(2309) & ChrW(2336) & ChrW(2340) & ChrW(2381) & ChrW(2340) & ChrW(2352), ChrW(2313) & ChrW(2344) & ChrW(2381) & ChrW(2344) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2309) & ChrW(2360) & ChrW(2367), ChrW(2319) & ChrW(2325) & ChrW(2366) & ChrW(2360) & ChrW(2368), _
						ChrW(2348) & ChrW(2351) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2330) & ChrW(2380) & ChrW(2352) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2346) & ChrW(2330) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2331) & ChrW(2351) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2360) & ChrW(2340) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2309) & ChrW(2336) & ChrW(2366) & ChrW(2360) & ChrW(2368), ChrW(2313) & ChrW(2344) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2319) & ChrW(2325) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), _
						ChrW(2348) & ChrW(2351) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2340) & ChrW(2367) & ChrW(2352) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2330) & ChrW(2380) & ChrW(2352) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2346) & ChrW(2334) & ChrW(2381) & ChrW(2330) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2331) & ChrW(2351) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2360) & ChrW(2344) & ChrW(2381) & ChrW(2340) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), _
						ChrW(2309) & ChrW(2344) & ChrW(2381) & ChrW(2336) & ChrW(2366) & ChrW(2344) & ChrW(2348) & ChrW(2381) & ChrW(2348) & ChrW(2375), ChrW(2313) & ChrW(2344) & ChrW(2344) & ChrW(2381) & ChrW(2360) & ChrW(2351))

	Specials = Array(ChrW(32) & ChrW(2360) & ChrW(2351) & ChrW(32), ChrW(32) & ChrW(2361) & ChrW(2332) & ChrW(2366) & ChrW(2352) & ChrW(32), ChrW(32) & ChrW(2354) & ChrW(2366) & ChrW(2326) & ChrW(32), ChrW(32) & ChrW(2325) & ChrW(2352) & ChrW(2379) & ChrW(2337) & ChrW(32), ChrW(32) & ChrW(2309) & ChrW(2352) & ChrW(2348) & ChrW(32), ChrW(32) & ChrW(2326) & ChrW(2352) & ChrW(2348) & ChrW(32))
    ConvertToNepaliText = GetNepali(Number)
End Function