VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Value As Long
Private m_IsPtr As Boolean

Friend Sub New_(ByVal v As Long)
    m_Value = v
    m_IsPtr = m_Value >= EReg.EReg_eax_
End Sub
Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Get IsPtr() As Boolean
    IsPtr = m_IsPtr
End Property


