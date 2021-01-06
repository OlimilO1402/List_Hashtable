VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Fill collection"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill hashtable"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "n:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HT As HashTable
Private m_SW As StopWatch

Private m_Col As Collection

Private Sub Command2_Click()
    'wieviel Speicher belegt dagegen eine Collection?
    Dim k As String
    Dim mess As String
    Dim i As Long
    Dim n As Long
    If IsNumeric(Text1.Text) Then
        n = CLng(Text1.Text)
    Else
        n = 100
    End If
    Set m_Col = New Collection
    m_SW.Reset
    m_SW.Start
    On Error Resume Next
    For i = 1 To n
        k = CStr(i * Rnd(CDbl(i)))
        'Call m_HT.Add(k, i)
        Call m_Col.Add(k, k)
    Next
    m_SW.SStop
    mess = "The collection now contains: " & CStr(m_Col.Count) & " elements. " & vbCrLf & _
           "with a capacity of: " & CStr("Capacity") & vbCrLf & _
           "Time needed to fill the collection: " & m_SW.ElapsedToString & vbCrLf & _
           "Show the hashtable in the listbox?"
           
    If MsgBox(mess, vbYesNo) = vbYes Then
        m_SW.Reset
        m_SW.Start
        List1.Visible = False
        'Call m_HT.ToListBox(List1)
        Call CollectionToListBox(m_Col, List1)
        List1.Visible = True
        m_SW.SStop
        mess = "Time needed to fill the listbox: " & m_SW.ElapsedToString
        MsgBox mess
    End If
    Label2.Caption = "Count: " & CStr(m_Col.Count)
    Label3.Caption = "Capacity: " & CStr("Capacity")
    Label4.Caption = CStr(m_HT.MemoryInBytes / 1024) & " kb"
    
End Sub

Public Sub CollectionToListBox(aCol As Collection, aLB As ListBox)
    Dim i   As Long
    Dim s   As String
    Dim p   As Long
    Dim cnt As Long 'Count 'Capacity
    
    cnt = aCol.Count 'CLng(CDbl(m_loadsize) / CDbl(m_loadFactor)) 'irgendwas is aber dann redundant oder?
    'oder auch einfacher:
    'cap = (UBound(m_buckets)+1)
    'bzw:
    'cap = Me.Capacity
    
    'Fleißaufgabe:
    'für die Stringlänge von key und val müßte man vorab das gesamte
    'Array einmal durchlaufen und die maximale Stringlänge ermitteln
    
    aLB.Clear
    p = Len(CStr(cnt))
    'Call aLB.AddItem("Count: " & CStr(m_Count) & " Capacity: " & CStr(cap))
    Dim v
    i = 0
    For Each v In aCol 'i = 0 To UBound(m_buckets)
        s = CStr(i) & ": " & PadLeft(CStr(v), 10)  'BucketToString(m_buckets(i))
        Call aLB.AddItem(s)
        i = i + 1
    Next
End Sub

Private Sub Form_Load()
    Set m_SW = New StopWatch
    Set m_HT = New HashTable
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call MString.DeleteStringPointers
End Sub
Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, h As Single
    l = List1.Left: T = List1.Top
    W = Me.ScaleWidth - l: h = Me.ScaleHeight - T
    If W > 0 And h > 0 Then Call List1.Move(l, T, W, h)
End Sub


Private Sub Command1_Click()
'
'schön wäre noch die Zeit anzuzeigen, wieviel Zeit es jeweils braucht
'A) die Hashtable zu füllen und
'B) die Hashtable in der ListBox anzuzeigen
'B könnte u.U viel mehr Zeit in Anspruch nehmen als A
'
    Dim k As String
    Dim mess As String
    Dim i As Long
    Dim n As Long
    If IsNumeric(Text1.Text) Then
        n = CLng(Text1.Text)
    Else
        n = 100
    End If
    
    Call m_HT.Clear
    m_SW.Reset
    m_SW.Start
    For i = 1 To n
        k = CStr(i * Rnd(i))
        Call m_HT.Add(k, i)
    Next
    m_SW.SStop
    mess = "The hashtable now contains: " & CStr(m_HT.Count) & " elements. " & vbCrLf & _
           "with a capacity of: " & CStr(m_HT.Capacity) & vbCrLf & _
           "Time needed to fill the hashtable: " & m_SW.ElapsedToString & vbCrLf & _
           "Show the hashtable in the listbox?"
           
    If MsgBox(mess, vbYesNo) = vbYes Then
        m_SW.Reset
        m_SW.Start
        List1.Visible = False
        Call m_HT.ToListBox(List1)
        List1.Visible = True
        m_SW.SStop
        mess = "Time needed to fill the listbox: " & m_SW.ElapsedToString
        MsgBox mess
    End If
    Label2.Caption = "Count: " & CStr(m_HT.Count)
    Label3.Caption = "Capacity: " & CStr(m_HT.Capacity)
    Label4.Caption = CStr(m_HT.MemoryInBytes / 1024) & " kb"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Command1_Click
    End If
End Sub
