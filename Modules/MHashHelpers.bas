Attribute VB_Name = "MHashHelpers"
Option Explicit
Public primes() As Long

Public Sub InitPrime()
   Call ArrayL(primes, 3&, 7&, 11&, &H11&, &H17&, &H1D&, &H25&, &H2F&, &H3B&, &H47&, &H59&, &H6B&, &H83&, &HA3&, &HC5&, &HEF&, _
   &H125&, &H161&, &H1AF&, &H209&, &H277&, &H2F9&, &H397&, &H44F&, &H52F&, &H63D&, &H78B&, &H91D&, &HAF1&, &HD2B&, &HFD1&, _
   &H12FD&, &H16CF&, &H1B65&, &H20E3&, &H2777&, &H2F6F&, &H38FF&, &H446F&, &H521F&, &H628D&, &H7655&, &H8E01&, &HAA6B&, _
   &HCC89&, &HF583&, &H126A7, &H1619B, &H1A857, &H1FD3B, &H26315, &H2DD67, &H3701B, &H42023, &H4F361, &H5F0ED, _
   &H72125, &H88E31, &HA443B, &HC51EB, &HEC8C1, &H11BDBF, &H154A3F, &H198C4F, &H1EA867, &H24CA19, &H2C25C1, _
   &H34FA1B, &H3F928F, &H4C4987, &H5B8B6F, &H6DDA89)
End Sub
Public Function ArrayL(ByRef arr() As Long, ParamArray params())
    ReDim arr(0 To UBound(params))
    Dim i As Long
    For i = 0 To UBound(params)
        arr(i) = CLng(params(i))
    Next
End Function
Public Function GetPrime(ByVal min As Long) As Long
    If (min < 0) Then
        'Throw New ArgumentException(Environment.GetResourceString("Arg_HTCapacityOverflow"))
        MsgBox "Arg_HTCapacityOverflow"
        Exit Function
    End If
    Dim i As Long
    Dim num2 As Long
    For i = 0 To UBound(primes)
        num2 = primes(i)
        If (num2 >= min) Then
            GetPrime = num2
            'Debug.Print "min: " & CStr(min) & " GetPrime 1 " & CStr(GetPrime)
            Exit Function
        End If
    Next i
    Dim j As Long: j = (min Or 1)
    Do While (j < &H7FFFFFFF)
        If MHashHelpers.IsPrime(j) Then
            GetPrime = j
            'Debug.Print "min: " & CStr(min) & " GetPrime 2 " & CStr(GetPrime)
            Exit Function
        End If
        j = (j + 2)
    Loop
    'Return min
    GetPrime = min
End Function
Public Function IsPrime(ByVal candidate As Long) As Boolean
    If ((candidate And 1) = 0) Then
        IsPrime = (candidate = 2)
        Exit Function
    End If
    Dim num As Long: num = CLng(VBA.Math.Sqr(CDbl(candidate)))
    Dim i As Long: i = 3
    Do While (i <= num)
        If ((candidate Mod i) = 0) Then
            IsPrime = False
            Exit Function
        End If
        i = (i + 2)
    Loop
    IsPrime = True
End Function
