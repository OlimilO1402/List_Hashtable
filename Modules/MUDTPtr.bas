Attribute VB_Name = "MUDTPtr"
Option Explicit

' Ein SafeArray-Descriptor dient in VB als ein universaler Zeiger
Public Type TUDTPtr
    pSA        As Long
    Reserved   As Long ' z.B. für IRecordInfo
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements  As Long
    lLBound    As Long
End Type

Public Enum SAFeature
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4

    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20
    FADF_HAVEIID = &H40
    FADF_HAVEVARTYPE = &H80

    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800
    FADF_RESERVED = &HF008
End Enum

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
                   ByRef pDst As Any, _
                   ByRef pSrc As Any, _
                   ByVal bLength As Long)

Public Declare Sub RtlZeroMemory Lib "kernel32" ( _
                   ByRef pDst As Any, _
                   ByVal bLength As Long)

Public Declare Function ArrPtr Lib "msvbvm60" _
                        Alias "VarPtr" ( _
                        ByRef pArr() As Any) As Long

Public Sub New_UDTPtr(ByRef this As TUDTPtr, _
                      ByVal Feature As SAFeature, _
                      ByVal bytesPerElement As Long, _
                      Optional ByVal CountElements As Long = 1, _
                      Optional ByVal lLBound As Long = 0)

    With this
        .pSA = VarPtr(.cDims)
        .cDims = 1
        .cbElements = bytesPerElement
        .fFeatures = CInt(Feature)
        .cElements = CountElements
        .lLBound = lLBound
    End With

End Sub

Public Sub AssignUDTPtr(pDst As TUDTPtr, pSrc As TUDTPtr)
'hier wird nicht einfach nur zugewiesen in der Art pDst = pSrc
'sondern hier wird nur pvdata und cElements zugewiesen, wobei
'cElements in Abhängigkeit von cbElement entsprechend angepasst wird
    pDst.pvData = pSrc.pvData
    If pDst.cbElements > 0 Then
        pDst.cElements = pSrc.cElements * pSrc.cbElements \ pDst.cbElements + 1
    End If
End Sub

Public Function UDTPtrToString(ByRef this As TUDTPtr) As String
' Um zu überprüfen ob der UDTPtr auch das enthält was er soll
' kann man folgende Funktion verwenden

    Dim s As String
    
    With this
        s = s & "pSA        : " & CStr(.pSA) & vbCrLf
        s = s & "Reserved   : " & CStr(.Reserved) & vbCrLf
        s = s & "cDims      : " & CStr(.cDims) & vbCrLf
        s = s & "fFeatures  : " & FeaturesToString(CLng(.fFeatures)) & vbCrLf
        s = s & "cbElements : " & CStr(.cbElements) & vbCrLf
        s = s & "cLocks     : " & CStr(.cLocks) & vbCrLf
        s = s & "pvData     : " & CStr(.pvData) & vbCrLf
        s = s & "cElements  : " & CStr(.cElements) & vbCrLf
        s = s & "lLBound    : " & CStr(.lLBound) & vbCrLf
    End With
    
    UDTPtrToString = s
    
End Function

Private Function FeaturesToString(ByVal f As SAFeature) As String
    
    Dim s As String
    Const sOr As String = " Or "
    
    If f And FADF_AUTO Then s = s & IIf(Len(s), sOr, "") & "FADF_AUTO"
    If f And FADF_STATIC Then s = s & IIf(Len(s), sOr, "") & "FADF_STATIC"
    If f And FADF_EMBEDDED Then s = s & IIf(Len(s), sOr, "") & "FADF_EMBEDDED"
    If f And FADF_FIXEDSIZE Then s = s & IIf(Len(s), sOr, "") & "FADF_FIXEDSIZE"
    If f And FADF_RECORD Then s = s & IIf(Len(s), sOr, "") & "FADF_RECORD"
    If f And FADF_HAVEIID Then s = s & IIf(Len(s), sOr, "") & "FADF_HAVEIID"
    If f And FADF_HAVEVARTYPE Then s = s & IIf(Len(s), sOr, "") & "FADF_HAVEVARTYPE"
    If f And FADF_BSTR Then s = s & IIf(Len(s), sOr, "") & "FADF_BSTR"
    If f And FADF_UNKNOWN Then s = s & IIf(Len(s), sOr, "") & "FADF_UNKNOWN"
    If f And FADF_DISPATCH Then s = s & IIf(Len(s), sOr, "") & "FADF_DISPATCH"
    If f And FADF_VARIANT Then s = s & IIf(Len(s), sOr, "") & "FADF_VARIANT"
    If f And FADF_RESERVED Then s = s & IIf(Len(s), sOr, "") & "FADF_RESERVED"
    
    FeaturesToString = s
    
End Function

