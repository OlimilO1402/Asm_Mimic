VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "r32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'class Register r32
Public Enum er32
    'Achtng, die Reihenfolge der Register hier im Enum nicht verändern!
    'Allzweckregister
    'Abk.   ' Name                ' Spezielle Verwendung
    eax_ = 0 ' Akkumulator         ' niederwertiger Teil und erster Operand bei Multiplikation; Quotient bei Division; BCD-Befehle
    ecx_     ' Zählregister        ' Schleifenzähler für JCXZ und LOOP; Verschiebedistanz von Verschiebebefehlen
    edx_     ' Datenregister       ' höherwerter Teil bei Multiplikation; Rest bei Division; Portadresse für die Befehle IN reg,DX und OUT DX,reg
    ebx_     ' Basisregister       ' keine spezielle Verwendung (Das 16-Bit-Register BX konnte im 16-Bit-Modus zur Index-Adressierung benutzt werden; im 32-Bit-Modus ist dies mit allen “General-Purpose”-Registern möglich.)
    
    'Indexregister
    esp_     ' Stackpointer        ' Zeiger auf die aktuelle Position im Stacksegment; nur eingeschränkt allgemein verwendbar, da dieses Register angibt, wo die Rücksprungadresse von Unterprogrammen und Interrupts gespeichert wird.
    ebp_     ' Basepointer         ' Wird zur Adressierung des Stackframe für lokale Variablen/Parameterübergabe benutzt.
    esi_     ' Source Index        ' Quelle für String-Operationen
    edi_     ' Destination Index   ' Ziel für String-Operationen

End Enum
Private m_Value As er32

Friend Sub New_(v As er32)
    m_Value = v
End Sub
Public Property Get Value() As er32
    Value = m_Value
End Property

