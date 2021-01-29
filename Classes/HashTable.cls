VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'mscorlib.System.Collections.HashTable
Private Type bucket
    Key As String 'Object
    val As String 'Object
    hash_coll As Long
End Type
Private m_buckets()  As bucket
Private m_Count      As Long
Private m_loadsize   As Long
Private m_loadFactor As Single
Private m_occupancy  As Long
Private m_Version    As Long
Private m_isWriterInProgress As Boolean

Private Sub Class_Initialize()
    'call InitHash
    'ReDim m_buckets(0 To 10)
    'Call rehash
    Call NewC(0, 1!)
End Sub
Public Sub NewC(ByVal Capacity As Long, ByVal loadFactor As Single)
    If (Capacity < 0) Then
        'Throw New ArgumentOutOfRangeException("capacity", Environment.GetResourceString("ArgumentOutOfRange_NeedNonNegNum"))
        MsgBox "ArgumentOutOfRangeException(capacity, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))"
    End If
    If ((loadFactor < 0.1!) Or (loadFactor > 1!)) Then
        'Throw New ArgumentOutOfRangeException("loadFactor", Environment.GetResourceString("ArgumentOutOfRange_HashtableLoadFactor", New Object() { 0.1, 1 }))
        MsgBox "ArgumentOutOfRangeException(loadFactor, Environment.GetResourceString(ArgumentOutOfRange_HashtableLoadFactor, New Object() { 0.1, 1 }))"
    End If
    m_loadFactor = (0.72! * loadFactor)
    Dim num As Double: num = (CSng(Capacity) / m_loadFactor)
    If (num > 2147483647) Then
        'Throw New ArgumentException(Environment.GetResourceString("Arg_HTCapacityOverflow")
        MsgBox "ArgumentException(Environment.GetResourceString(Arg_HTCapacityOverflow)"
    End If
    Dim num2 As Long: num2 = IIf((num > 11), MHashHelpers.GetPrime(CLng(num)), 11)
    ReDim m_buckets(0 To num2 - 1)
    m_loadsize = CLng(m_loadFactor * num2)
    m_isWriterInProgress = False
End Sub

Public Sub Clear()
    m_isWriterInProgress = True
    ReDim m_buckets(0 To 10)
    m_loadsize = CLng(0.72 * 11)
    m_Count = 0
    m_occupancy = 0
    m_Version = m_Version + 1
    m_isWriterInProgress = False
End Sub
Private Function InitHash(ByVal Key As String, ByVal hashsize As Long, ByRef seed As Long, ByRef incr As Long) As Long
    'Dim num As Long: num = (Me.GetHash(key) And &H7FFFFFFF)
    InitHash = (MString.GetHashCode(Key) And &H7FFFFFFF)
    seed = InitHash 'num
    incr = (1 + ((((seed And &HFFFFFFE0) \ &H20&) + 1) Mod (hashsize - 1)))
    'InitHash = num
End Function
 
Public Sub Add(ByVal Key As String, ByVal val As String)
    Call Me.Insert(Key, val, True)
End Sub
Public Property Get Count() As Long
    Count = m_Count
End Property
Public Property Get Capacity() As Long
    'Capacity = UBound(m_buckets) + 1
    Capacity = buckets_Length
End Property
Public Sub Insert(Key As String, nvalue As String, ByVal bAdd As Boolean)
    Dim num As Long
    Dim num2 As Long
    If LenB(Key) = 0 Then
        MsgBox "key, Environment.GetResourceString(ArgumentNull_Key)"
    End If
    If (Me.Count >= m_loadsize) Then
        Call expand
    ElseIf ((m_occupancy > m_loadsize) And (m_Count > 100)) Then
        Call rehash
    End If
    Dim num3 As Long: num3 = InitHash(Key, buckets_Length, num, num2) 'num und num2 werden zurückgegeben
    Dim num4 As Long ' = 0
    Dim index As Long: index = -1
    Dim num6 As Long: num6 = CLng((num Mod (buckets_Length)))
Label_0071: ''häh??? key is array?
    'If (((index = -1) And (m_buckets(num6).key Is m_buckets)) And (m_buckets(num6).hash_coll < 0)) Then
    'If (index = -1) And (m_buckets(num6).hash_coll < 0) Then
    '    index = num6
    'End If
    'If ((Len(m_buckets(num6).key) = 0) Or ((m_buckets(num6).key Is Me.buckets) And ((m_buckets(num6).hash_coll And &H80000000) = 0))) Then
    If (((Len(m_buckets(num6).Key) = 0) And ((m_buckets(num6).hash_coll And &H80000000) = 0))) Then
        If (index <> -1) Then
            num6 = index
        End If
        'Thread.BeginCriticalRegion
        m_isWriterInProgress = True
        m_buckets(num6).val = nvalue
        m_buckets(num6).Key = Key
        m_buckets(num6).hash_coll = (m_buckets(num6).hash_coll Or CLng(num3))
        m_Count = m_Count + 1
        Call UpdateVersion
        m_isWriterInProgress = False
        'Thread.EndCriticalRegion
    ElseIf (((m_buckets(num6).hash_coll And &H7FFFFFFF) = num3) And (m_buckets(num6).Key = Key)) Then
        If bAdd Then
            'Throw New ArgumentException(Environment.GetResourceString("Argument_AddingDuplicate__", New Object() { Me.buckets(num6).key, key }))
            MsgBox "ArgumentException(Environment.GetResourceString(Argument_AddingDuplicate__, New Object() { Me.buckets(num6).key, key }))"
        End If
        'Thread.BeginCriticalRegion
        m_isWriterInProgress = True
        m_buckets(num6).val = nvalue
        'Me.UpdateVersion
        m_Version = m_Version + 1
        m_isWriterInProgress = False
        'Thread.EndCriticalRegion
    Else
        If ((index = -1) And (m_buckets(num6).hash_coll >= 0)) Then
            m_buckets(num6).hash_coll = (m_buckets(num6).hash_coll Or -2147483648#)
            m_occupancy = m_occupancy + 1
        End If
        num6 = CLng(((num6 + num2) Mod CLng(UBound(m_buckets) + 1)))
        If (num4 < (UBound(m_buckets) + 1)) Then
            GoTo Label_0071
        End If
        If (index = -1) Then
            'Throw New InvalidOperationException(Environment.GetResourceString("InvalidOperation_HashInsertFailed"))
            MsgBox "InvalidOperationException(Environment.GetResourceString(InvalidOperation_HashInsertFailed))"
        End If
        'Thread.BeginCriticalRegion
        m_isWriterInProgress = True
        m_buckets(index).val = nvalue
        m_buckets(index).Key = Key
        m_buckets(index).hash_coll = (m_buckets(index).hash_coll Or CLng(num3))
        m_Count = m_Count + 1
        'Me.UpdateVersion
        m_Version = m_Version + 1
        m_isWriterInProgress = False
        'Thread.EndCriticalRegion
    End If

End Sub
Private Sub expand()
    Dim prime As Long: prime = MHashHelpers.GetPrime(buckets_Length * 2)
    Call rehash(prime)
End Sub
Private Function buckets_Length() As Long
    If SAPtr(ArrPtr(m_buckets)) <> 0 Then
        buckets_Length = UBound(m_buckets) + 1
    Else
        buckets_Length = 0
    End If
End Function
Private Sub UpdateVersion()
    m_Version = m_Version + 1
End Sub
Private Sub rehash(Optional ByVal newsize As Long)
    If newsize = 0 Then newsize = buckets_Length 'UBound(m_buckets) + 1
    m_occupancy = 0
    'ReDim Preserve m_buckets(0 To newsize - 1)
    ReDim newBuckets(0 To newsize - 1) As bucket
    'Dim newBuckets As bucket() = New bucket(newsize  - 1) {}
    Dim i As Long
    Dim bucket As bucket
    For i = 0 To buckets_Length - 1 '.Length - 1
        bucket = m_buckets(i)
        If Not (Len(bucket.Key) = 0) Then 'And (Not bucket.key Is Me.buckets)) Then
            Call putEntry(newBuckets, bucket.Key, bucket.val, (bucket.hash_coll And &H7FFFFFFF))
        'Else
        '    Debug.Print "shit schon da!"
        End If
    Next i
    'Thread.BeginCriticalRegion
    m_isWriterInProgress = True
    m_buckets = newBuckets
    m_loadsize = CLng(m_loadFactor * newsize)
    Call UpdateVersion
    m_isWriterInProgress = False
    'Thread.EndCriticalRegion
End Sub
Private Sub putEntry(ByRef newBuckets() As bucket, Key As String, nvalue As String, ByVal hashcode As Long)
    Dim newBuckets_Length As Long: newBuckets_Length = UBound(newBuckets) + 1
    Dim num As Long: num = hashcode 'DirectCast(hashcode, UInt32)
    Dim num2 As Long:   num2 = (1 + ((((num And &HFFFFFFE0) \ &H20&) + 1) Mod (newBuckets_Length - 1)))
    Dim index As Long: index = CLng((num Mod (newBuckets_Length)))
Label_0017:
    If Len(newBuckets(index).Key) = 0 Then 'Or (newBuckets(index).key Is Me.buckets)) Then
        'vielleicht wäre es besser die strings rüberzuswappen
        newBuckets(index).val = nvalue
        newBuckets(index).Key = Key
        newBuckets(index).hash_coll = (newBuckets(index).hash_coll Or hashcode)
    Else
        If (newBuckets(index).hash_coll >= 0) Then
            newBuckets(index).hash_coll = (newBuckets(index).hash_coll Or -2147483648#)
            m_occupancy = m_occupancy + 1
        End If
        index = CLng(((index + num2) Mod CLng(newBuckets_Length)))
        GoTo Label_0017
    End If
End Sub

Public Function Contains(Key As String) As Boolean
    Contains = Me.ContainsKey(Key)
End Function
Public Function ContainsKey(Key As String) As Boolean
    Dim num As Long
    Dim num2 As Long
    'Dim bucket As bucket
    If LenB(Key) = 0 Then
        'Throw New ArgumentNullException("key", Environment.GetResourceString("ArgumentNull_Key"))
        MsgBox "ArgumentNullException(key, Environment.GetResourceString(ArgumentNull_Key))"
        Exit Function
    End If
    'Dim buckets As bucket() = Me.buckets
    Dim num3 As Long: num3 = InitHash(Key, UBound(m_buckets) + 1, num, num2)
    Dim num4 As Long ' = 0
    Dim index As Long: index = CLng((num Mod (UBound(m_buckets) + 1)))
    Do
        'bucket = buckets(index)
        If LenB(m_buckets(index).Key) = 0 Then
            'Return False
            ContainsKey = False
            Exit Function
        End If
        If (((m_buckets(index).hash_coll And &H7FFFFFFF) = num3) And (m_buckets(index).Key = Key)) Then
            'Return True
            ContainsKey = True
            Exit Function
        End If
        index = CLng(((index + num2) Mod CLng(UBound(m_buckets) + 1)))
    Loop While ((m_buckets(index).hash_coll < 0) And (Inc(num4) < (UBound(m_buckets) + 1)))
    'Return False
    ContainsKey = False
End Function
Public Function ContainsValue(Value As String) As Boolean
    If LenB(Value) = 0 Then
        Dim length As Long: length = UBound(m_buckets) + 1
        Do While length >= 0
            length = length - 1
            If (((Not (LenB(m_buckets(length).Key) = 0)) And (Not (LenB(m_buckets(length).val) = 0)))) Then 'And (Not m_buckets(length).key Is Me.buckets))
                ContainsValue = True
                Exit Function
            End If
        Loop
    Else
        Dim index As Long: index = UBound(m_buckets) + 1
        Do While index >= 0
            index = index - 1
            'Dim val As Object = M_buckets(index).val
            If Not (LenB(m_buckets(index).val) = 0) Then
                If StrComp(m_buckets(index).val, Value, vbBinaryCompare) = 0 Then
                    ContainsValue = True
                    Exit Function
                End If
            End If
        Loop
    End If
    ContainsValue = False
End Function

' Default
Public Property Get Item(Key As String) As String
    Dim num As Long
    Dim num2 As Long
    Dim bucket As bucket
    Dim version As Long
    Dim num7 As Long
    If LenB(Key) = 0 Then
        'Throw New ArgumentNullException("key", Environment.GetResourceString("ArgumentNull_Key"))
        MsgBox "ArgumentNullException(key, Environment.GetResourceString(ArgumentNull_Key))"
    End If
    'Dim buckets As bucket() = Me.buckets
    Dim num3 As Long: num3 = InitHash(Key, buckets_Length, num, num2)
    Dim num4 As Long ' = 0
    Dim index As Long: index = CLng((num Mod buckets_Length))
Label_0038:
    num7 = 0
    Do
        version = m_Version
        bucket = m_buckets(index)
        num7 = num7 + 1
        If ((num7 Mod 8) = 0) Then
            '???? wieso das denn?
            'Thread.Sleep (1)
        End If
    Loop While (m_isWriterInProgress Or (Not (version = m_Version)))
    If (Not (LenB(bucket.Key) = 0)) Then
        If (((bucket.hash_coll And &H7FFFFFFF) = num3) And (StrComp(bucket.Key, Key, vbBinaryCompare) = 0)) Then
            Item = bucket.val
            Exit Property
        End If
        index = CLng(((index + num2) Mod buckets_Length))
        num4 = num4 + 1
        If ((bucket.hash_coll < 0) And (num4 < (buckets_Length))) Then
            GoTo Label_0038
        End If
    End If
    'Return Nothing
End Property
Public Property Let Item(Key As String, Value As String)
    Call Me.Insert(Key, Value, False)
End Property
Public Property Get index(Key As String) As Long
    Dim num As Long
    Dim num2 As Long
    Dim bucket As bucket
    Dim version As Long
    Dim num7 As Long
    If LenB(Key) = 0 Then
        'Throw New ArgumentNullException("key", Environment.GetResourceString("ArgumentNull_Key"))
        MsgBox "ArgumentNullException(key, Environment.GetResourceString(ArgumentNull_Key))"
    End If
    'Dim buckets As bucket() = Me.buckets
    Dim num3 As Long: num3 = InitHash(Key, buckets_Length, num, num2)
    Dim num4 As Long ' = 0
    'Dim Index As Long:
    index = CLng((num Mod (buckets_Length)))
Label_0038:
    num7 = 0
    Do
        version = m_Version
        bucket = m_buckets(index)
        num7 = num7 + 1
        If ((num7 Mod 8) = 0) Then
            '???? wieso das denn?
            'Thread.Sleep (1)
        End If
    Loop While (m_isWriterInProgress Or (Not (version = m_Version)))
    If (Not (LenB(bucket.Key) = 0)) Then
        If (((bucket.hash_coll And &H7FFFFFFF) = num3) And (StrComp(bucket.Key, Key, vbBinaryCompare) = 0)) Then
            'Item = bucket.val
            Exit Property
        End If
        index = CLng(((index + num2) Mod CLng(buckets_Length)))
        num4 = num4 + 1
        If ((bucket.hash_coll < 0) And (num4 < (buckets_Length))) Then
            GoTo Label_0038
        End If
    End If
    'Return Nothing
End Property

Public Property Get Key(index As Long) As String
    Key = m_buckets(index).Key
End Property
Public Property Get Value(index As Long) As String
    Value = m_buckets(index).val
End Property
Public Property Get HashColl(index As Long) As Long
    HashColl = m_buckets(index).hash_coll
End Property

Private Function BucketToString(b As bucket) As String
    Dim s As String
    With b
        s = s & PadLeft(.Key, 10) & ", "
        s = s & PadLeft(.val, 10) & ", "
        s = s & Hex$(.hash_coll)
    End With
    BucketToString = s
End Function
Public Sub ToListBox(aLB As ListBox)
    Dim i   As Long
    Dim s   As String
    Dim p   As Long
    Dim cap As Long 'Capacity
    
    cap = CLng(CDbl(m_loadsize) / CDbl(m_loadFactor)) 'irgendwas is aber dann redundant oder?
    'oder auch einfacher:
    'cap = (UBound(m_buckets)+1)
    'bzw:
    'cap = Me.Capacity
    
    'Fleißaufgabe:
    'für die Stringlänge von key und val müßte man vorab das gesamte
    'Array einmal durchlaufen und die maximale Stringlänge ermitteln
    
    aLB.Clear
    p = Len(CStr(cap))
    'Call aLB.AddItem("Count: " & CStr(m_Count) & " Capacity: " & CStr(cap))
    For i = 0 To UBound(m_buckets)
        s = PadLeft(CStr(i), p) & ": " & BucketToString(m_buckets(i))
        Call aLB.AddItem(s)
    Next
End Sub

Public Property Get MemoryInBytes() As Long
    'liefert zurück wieviel Speicher durch die Hashtable belegt wird
    Dim i As Long, b As Long, l As Long
    Dim cap As Long: cap = Capacity
    For i = 0 To cap - 1
        b = b + 12
        With m_buckets(i)
            l = LenB(.Key)
            If l > 0 Then b = b + l + 4
            l = LenB(.val)
            If l > 0 Then b = b + l + 4
        End With
    Next
    MemoryInBytes = b
End Property
