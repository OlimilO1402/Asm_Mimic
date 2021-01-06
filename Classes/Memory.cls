VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr()
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 2 ^ 10 - 1)
End Sub

Public Sub Add(ByVal Value As Long)
    Dim u As Long: u = UBound(m_Arr)
    If u < m_Count Then
        MsgBox "Speicher voll!"
        'ReDim Preserve m_Arr(0 To 2 * m_Count - 1)
        Exit Sub
    End If
    m_Arr(m_Count) = Value
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Let Item(ByVal Index As Long, Value As Variant)
    Dim u As Long: u = UBound(m_Arr)
    If u < Index Then
        MsgBox "Index zu groß!"
        Exit Property
    End If
    If IsObject(Value) Then
        Set m_Arr(Index) = Value
    Else
        m_Arr(Index) = Value
    End If
End Property
Public Property Get Item(ByVal Index As Long) As Variant
    Dim u As Long: u = UBound(m_Arr)
    If u < Index Then
        MsgBox "Index zu groß!"
        Exit Property
    End If
    If IsObject(m_Arr(Index)) Then
        Set Item = m_Arr(Index)
    Else
        Item = m_Arr(Index)
    End If
End Property
