VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RegEAX As Long
Public RegECX As Long
Public RegEDX As Long
Public RegEBX As Long
Public RegESP As Long
Public RegEBP As Long
Public RegESI As Long
Public RegEDI As Long

Public Property Get Register(reg As r32) As Long
    Select Case reg.Value
    Case er32.eax_: Register = Me.RegEAX
    Case er32.ecx_: Register = Me.RegECX
    Case er32.edx_: Register = Me.RegEDX
    Case er32.ebx_: Register = Me.RegEBX
    Case er32.esp_: Register = Me.RegESP
    Case er32.ebp_: Register = Me.RegEBP
    Case er32.esi_: Register = Me.RegESI
    Case er32.edi_: Register = Me.RegEDI
    End Select
End Property

Public Property Let Register(reg As r32, Value As Long)
    Select Case reg.Value
    Case er32.eax_: Me.RegEAX = Value
    Case er32.ecx_: Me.RegECX = Value
    Case er32.edx_: Me.RegEDX = Value
    Case er32.ebx_: Me.RegEBX = Value
    Case er32.esp_: Me.RegESP = Value
    Case er32.ebp_: Me.RegEBP = Value
    Case er32.esi_: Me.RegESI = Value
    Case er32.edi_: Me.RegEDI = Value
    End Select
End Property
