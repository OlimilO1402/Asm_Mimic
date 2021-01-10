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
' Vol. 1 S. 3-10 Kap. 3.4 BASIC PROGRAM EXECUTION REGISTERS
' Vol. 1 S. 3-12 Figure 3-5.  Alternate General-Purpose Register Names
'The 386 32-bit Registers
'General Purpose Registers
'| 31            16|15     8|7      0|  16Bit   32Bit
'|                 |   AH   |   AL   |   AX      EAX
'|                 |   CH   |   CL   |   CX      ECX
'|                 |   DH   |   DL   |   DX      EDX
'|                 |   BH   |   BL   |   BX      EBX
'
'|                 |       BP        |           EBP
'|                 |       SI        |           ESI
'|                 |       DI        |           EDI
'|                 |       SP        |           ESP
Public Enum EReg
    EReg_al = &H0&:    EReg_cl = &H1&:    EReg_dl = &H2&:    EReg_bl = &H3&
    
    EReg_ah = &H4&:    EReg_ch = &H5&:    EReg_dh = &H6&:    EReg_bh = &H7&
    
    EReg_ax = &H10&:   EReg_cx = &H11&:   EReg_dx = &H12&:   EReg_bx = &H13&
    
    EReg_sp = &H14&:   EReg_bp = &H15&:   EReg_si = &H16&:   EReg_di = &H17&
    
    'Allzweckregister
    'Abk.      ' Name                ' Spezielle Verwendung
    EReg_eax = &H20& ' Akkumulator         ' niederwertiger Teil und erster Operand bei Multiplikation; Quotient bei Division; BCD-Befehle
    EReg_ecx = &H21& ' Zählregister        ' Schleifenzähler für JCXZ und LOOP; Verschiebedistanz von Verschiebebefehlen
    EReg_edx = &H22& ' Datenregister       ' höherwerter Teil bei Multiplikation; Rest bei Division; Portadresse für die Befehle IN reg,DX und OUT DX,reg
    EReg_ebx = &H23& ' Basisregister       ' keine spezielle Verwendung (Das 16-Bit-Register BX konnte im 16-Bit-Modus zur Index-Adressierung benutzt werden; im 32-Bit-Modus ist dies mit allen “General-Purpose”-Registern möglich.)
    
    'Indexregister
    EReg_esp = &H24& ' Stackpointer        ' Zeiger auf die aktuelle Position im Stacksegment; nur eingeschränkt allgemein verwendbar, da dieses Register angibt, wo die Rücksprungadresse von Unterprogrammen und Interrupts gespeichert wird.
    EReg_ebp = &H25& ' Basepointer         ' Wird zur Adressierung des Stackframe für lokale Variablen/Parameterübergabe benutzt.
    EReg_esi = &H26& ' Source Index        ' Quelle für String-Operationen
    EReg_edi = &H27& ' Destination Index   ' Ziel für String-Operationen
    
    EReg_eax_ = &H40& ' [eax] contains Pointer
    EReg_ecx_ = &H41& ' [ecx] contains Pointer
    EReg_edx_ = &H42& ' [edx] contains Pointer
    EReg_ebx_ = &H43& ' [ebx] contains Pointer
    
    EReg_ebp_ = &H45& ' [ebp] contains Pointer
    EReg_esi_ = &H46& ' [esi] contains Pointer
    EReg_edi_ = &H47& ' [edi] contains Pointer
End Enum

Private Type TRegLH8H16
    lo8 As Byte
    hi8 As Byte
    hi16 As Integer
End Type
Private Type TRegLH16
    lo16 As Integer
    hi16 As Integer
End Type
Private Type TReg32
    Value As Long
End Type

Private m_Reg_AL_AH_AX_EAX As TReg32 'these fields contain all the values in the registers
Private m_Reg_CL_CH_CX_ECX As TReg32
Private m_Reg_DL_DH_DX_EDX As TReg32
Private m_Reg_BL_BH_BX_EBX As TReg32
Private m_Reg_SP_ESP As TReg32
Private m_Reg_BP_EBP As TReg32
Private m_Reg_SI_ESI As TReg32
Private m_Reg_DI_EDI As TReg32

Public Property Let Register(areg As reg, ByVal Value As Variant)
    Dim reg As TReg32
    Dim v As Long: v = areg.Value
    If 3 < v And v <= 7 Then v = v - 4
    Select Case v And &HF&
    Case 0: reg = m_Reg_AL_AH_AX_EAX
    Case 1: reg = m_Reg_CL_CH_CX_ECX
    Case 2: reg = m_Reg_DL_DH_DX_EDX
    Case 3: reg = m_Reg_BL_BH_BX_EBX
    Case 4: reg = m_Reg_SP_ESP
    Case 5: reg = m_Reg_BP_EBP
    Case 6: reg = m_Reg_SI_ESI
    Case 7: reg = m_Reg_DI_EDI
    End Select
    If areg.Value < &H4& Then
        SetL8 reg, Value
    ElseIf areg.Value < &H10& Then
        SetH8 reg, Value
    ElseIf areg.Value < &H20& Then
        SetL16 reg, Value
    Else
        reg.Value = CLng(Value)
    End If
    Select Case v And &HF&
    Case 0: m_Reg_AL_AH_AX_EAX = reg
    Case 1: m_Reg_CL_CH_CX_ECX = reg
    Case 2: m_Reg_DL_DH_DX_EDX = reg
    Case 3: m_Reg_BL_BH_BX_EBX = reg
    Case 4: m_Reg_SP_ESP = reg
    Case 5: m_Reg_BP_EBP = reg
    Case 6: m_Reg_SI_ESI = reg
    Case 7: m_Reg_DI_EDI = reg
    End Select
End Property
Private Sub SetL8(rDst As TReg32, Value As Variant)
    Dim rSrc As TRegLH8H16:  LSet rSrc = rDst 'zuerst so rüber kopieren, um die anderen Werte zu sichern, dann die 8-bit austauschen und wieder zurückkopieren
    rSrc.lo8 = CByte(Value): LSet rDst = rSrc
End Sub
Private Sub SetH8(rDst As TReg32, Value As Variant)
    Dim rSrc As TRegLH8H16:  LSet rSrc = rDst
    rSrc.hi8 = CByte(Value): LSet rDst = rSrc
End Sub
Private Sub SetL16(rDst As TReg32, Value As Variant)
    Dim rSrc As TRegLH16:    LSet rSrc = rDst
    rSrc.lo16 = CInt(Value): LSet rDst = rSrc
End Sub

Public Property Get Register(areg As reg) As Variant 'Byte Or Integer Or Long
    Dim reg As TReg32
    Dim v As Long: v = areg.Value
    If 3 < v And v <= 7 Then v = v - 4
    Select Case v And &HF& 'das Register rausfiltern
    Case 0: reg = m_Reg_AL_AH_AX_EAX
    Case 1: reg = m_Reg_CL_CH_CX_ECX
    Case 2: reg = m_Reg_DL_DH_DX_EDX
    Case 3: reg = m_Reg_BL_BH_BX_EBX
    Case 4: reg = m_Reg_SP_ESP
    Case 5: reg = m_Reg_BP_EBP
    Case 6: reg = m_Reg_SI_ESI
    Case 7: reg = m_Reg_DI_EDI
    End Select
    If areg.Value < &H4 Then
        Register = GetL8(reg)
    ElseIf areg.Value < &H10& Then
        Register = GetH8(reg)
    ElseIf areg.Value < &H20& Then
        Register = GetL16(reg)
    Else
        Register = reg.Value
    End If
End Property
Private Function GetL8(rSrc As TReg32) As Byte
    Dim rDst As TRegLH8H16: LSet rDst = rSrc: GetL8 = rDst.lo8
End Function
Private Function GetH8(rSrc As TReg32) As Byte
    Dim rDst As TRegLH8H16: LSet rDst = rSrc: GetH8 = rDst.hi8
End Function
Private Function GetL16(rSrc As TReg32) As Integer
    Dim rDst As TRegLH16: LSet rDst = rSrc: GetL16 = rDst.lo16
End Function
