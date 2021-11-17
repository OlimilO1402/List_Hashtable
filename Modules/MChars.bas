Attribute VB_Name = "MChars"
Option Explicit
Public Type TCharPointer
    pudt    As TUDTPtr
    Chars() As Integer
End Type
Public Type TLongPointer
    pudt     As TUDTPtr
    Values() As Long
End Type

Private str  As TCharPointer
Private strL As TLongPointer
Public Type TCur
    c As Currency
End Type
Public Type TLngHiLo
    Hi As Long
    Lo As Long
End Type

Public Sub InitMString()
    Call New_CharPointer(str, "")
    Call New_LongPointer(strL, 0) 'wird nur für GetHashCode gebraucht
End Sub
Public Sub DeleteStringPointers()
    Call DeleteCharPointer(str)
    Call DeleteLongPointer(strL)
End Sub
Public Sub New_CharPointer(ByRef this As TCharPointer, ByRef StrVal As String)
    With this
        Call New_UDTPtr(.pudt, FADF_AUTO Or FADF_FIXEDSIZE, 2, Len(StrVal), 1)
        With .pudt
            .pvData = StrPtr(StrVal)
        End With
        Call RtlMoveMemory(ByVal ArrPtr(.Chars), ByVal VarPtr(.pudt), 4)
    End With
End Sub
Public Sub DeleteCharPointer(ByRef this As TCharPointer)
    With this
        Call RtlZeroMemory(ByVal ArrPtr(.Chars), 4)
    End With
End Sub
Public Sub New_LongPointer(ByRef this As TLongPointer, ByVal pLong As Long)
    With this
        Call New_UDTPtr(.pudt, FADF_AUTO Or FADF_FIXEDSIZE, 4)
        With .pudt
            .pvData = pLong
        End With
        Call RtlMoveMemory(ByVal ArrPtr(.Values), ByVal VarPtr(.pudt), 4)
    End With
End Sub
Public Sub DeleteLongPointer(ByRef this As TLongPointer)
    With this
        Call RtlZeroMemory(ByVal ArrPtr(.Values), 4)
    End With
End Sub


Private Function InitStrPtr(this As String)
    With str.pudt
        .pvData = StrPtr(this)
        .cElements = Len(this)
    End With
End Function

Public Function GetHashCode(this As String) As Long
    Call InitStrPtr(this)
    'funzt nur mit:
    'Projekt -> Eigenschaften -> Kompilieren -> Weitere Optimierungen -> keine Überprüfung auf Ganzzahlüberlauf
    'und dann auch nur kompiliert!
    Call MUDTPtr.AssignUDTPtr(strL.pudt, str.pudt)
    Dim num1 As Long: num1 = &H15051505
    Dim num2 As Long: num2 = num1
    Dim numPtr As Long
    Dim ShL5_num1 As Long, ShR27_num1 As Long
    Dim ShL5_num2 As Long, ShR27_num2 As Long
    Dim i As Long: i = Len(this)
    Do While i > 0             ' nicht ShRZ27
        
        'Inlining von ShL5(num1)
        If num1 And &H4000000 Then
            ShL5_num1 = (num1 And &H3FFFFFF) * &H20& Or &H80000000
        Else
            ShL5_num1 = (num1 And &H3FFFFFF) * &H20&
        End If
        
        'Inlining von ShR27(num1)
        ShR27_num1 = (num1 And &HF8000000) \ &H8000000
        
        'nur für die Native-exe ist die Addition oder Multiplikation sicher
        #If IsNative Then
            num1 = (ShL5_num1 + num1 + ShR27_num1) Xor strL.Values(numPtr)
        #Else
            num1 = UAddC(UAddC(ShL5_num1, num1), ShR27_num1) Xor strL.Values(numPtr)
        #End If
        
        If (i <= 2) Then
            Exit Do
        End If
        
        'Inlining von ShL5(num2)
        If num2 And &H4000000 Then
            ShL5_num2 = (num2 And &H3FFFFFF) * &H20& Or &H80000000
        Else
            ShL5_num2 = (num2 And &H3FFFFFF) * &H20&
        End If
        
        'Inlining von ShR27(num2)
        ShR27_num2 = (num2 And &HF8000000) \ &H8000000
        
        'nur für die Native-exe ist die Addition oder Multiplikation sicher
        #If IsNative Then
            num2 = (ShL5_num2 + num2 + ShR27_num2) Xor strL.Values(numPtr + 1)
        #Else
            num2 = UAddC(UAddC(ShL5_num2, num2), ShR27_num2) Xor strL.Values(numPtr + 1)
        #End If
        
        numPtr = numPtr + 2
        i = i - 4
    Loop
    #If IsNative Then
        GetHashCode = num1 + num2 * &H5D588B65
    #Else
        GetHashCode = UAddC(num1, MulOFlow(num2, &H5D588B65))
    #End If
    
End Function
'die Funktionen Multiplikation und Addition IDE-safe machen:
'eine vergleichbare Funktion mit RtlMoveMemory wäre hier
'_nicht_ schneller sondern ca 20-30% langsamer
'mit GetMem4/PutMem4 dagegen nochmal 10-20% schneller
Public Function MulOFlow(ByVal a As Long, ByVal b As Long) As Long
    'führt eine überlaufsichere unsigned Multiplikation mit zwei signed Int32 durch
    'Gibt die unteren 4-Byte eines Int64 bei einer Multiplitkation zurück,
    'selbst wenn ein Int32-Overflow stattfinden würde.
    Dim al As TLngHiLo, ac As TCur
    al.Hi = a:       LSet ac = al
    ac.c = ac.c * b: LSet al = ac
    MulOFlow = al.Hi
End Function

Public Function UAddC(ByVal a As Long, ByVal b As Long) As Long
    'führt eine überlaufsichere unsigned Addition mit zwei signed Int32 durch
    '-> entspricht der gleichen Bitfolge wie bei unsigned Int32,
    'bei einem Überlauf wird nur der untere 4-Byte Anteil zurückgegeben
    Dim ll As TLngHiLo, ac As TCur, bc As TCur
    ll.Hi = a: LSet ac = ll
    ll.Hi = b: LSet bc = ll
    ac.c = ac.c + bc.c
    LSet ll = ac: UAddC = ll.Hi
End Function
'
'Public Function PadLeft(this As String, _
'                        ByVal totalWidth As Long, _
'                        Optional ByVal paddingChar As String) As String
'    If LenB(paddingChar) Then
'        If Len(this) < totalWidth Then
'            PadLeft = String$(totalWidth, paddingChar)
'            MidB$(PadLeft, totalWidth * 2 - LenB(this) + 1) = this
'        Else
'            PadLeft = this
'        End If
'    Else
'        PadLeft = Space$(totalWidth)
'        RSet PadLeft = this
'    End If
'End Function
'Public Function PadRight(this As String, _
'                         ByVal totalWidth As Long, _
'                         Optional ByVal paddingChar As String) As String
'    If LenB(paddingChar) Then
'        If Len(this) < totalWidth Then
'            PadRight = String$(totalWidth, paddingChar)
'            MidB$(PadRight, 1) = this
'        Else
'            PadRight = this
'        End If
'    Else
'        PadRight = Space$(totalWidth)
'        LSet PadRight = this
'    End If
'End Function
'
'

