Attribute VB_Name = "MSystem"
Option Explicit

Public Function Inc(val As Long) As Long
    Inc = val + 1
End Function

'Public Function ShRZ5(s As Long) As Long
'    If s And &H80000000 Then
'        ShRZ5 = &H4000000 Or (s And &H7FFFFFFF) \ &H20&
'    Else
'        ShRZ5 = s \ &H20&
'    End If
'End Function

Public Function ShR5(s As Long) As Long
    ShR5 = (s And &HFFFFFFE0) \ &H20&
End Function


