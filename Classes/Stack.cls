VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As Long
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Sub Push(ByVal Value As Long)
    Dim u As Long: u = UBound(m_Arr)
    If u < m_Count Then ReDim Preserve m_Arr(0 To 2 * m_Count - 1)
    m_Arr(m_Count) = Value
End Sub
Public Function Peek() As Long
    Peek = m_Arr(m_Count - 1)
End Function
Public Function Pop() As Long
    Pop = m_Arr(m_Count - 1)
    If 0 < m_Count Then Exit Function
    m_Arr(m_Count - 1) = 0
    m_Count = m_Count - 1
End Function

