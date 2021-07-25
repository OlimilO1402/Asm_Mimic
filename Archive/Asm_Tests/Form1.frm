VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnAsmAdd 
      Caption         =   "AsmAdd"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal bytLen As Long)

Private Declare Function CallWindowProcA Lib "user32" ( _
                 ByVal pFnc As Long, _
                 ByVal v1 As Long, _
                 ByVal v2 As Long, _
                 ByVal v3 As Long, _
                 ByVal v4 As Long) As Long

Private m_VMemPtr As Long
Private m_MemSize As Long

Private Sub Form_Load()
    m_MemSize = 128
    m_VMemPtr = MVMem.VMemPtr(m_MemSize)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MVMem.VMemPtrFree m_VMemPtr
End Sub

Private Sub BtnAsmAdd_Click()

    ReDim bytes(0 To m_MemSize - 1) As Byte
    Dim sum1 As Long: sum1 = 12345
    Dim sum2 As Long: sum2 = 23456
    Dim sum  As Long ' sum = sum1 + sum2 = 12345 + 23456 = 35801
    Dim i As Long, c As Long
    
    'pop eax
    'pop ecx
    'add eax, ecx
    'ret eax
    bytes(i) = &H58: i = i + 1
    bytes(i) = &H59: i = i + 1
    bytes(i) = &H59: i = i + 1
    
    'c = LenB(sum1)
    'memcpy bytes(i), sum1, c: i = i + c
    
    sum = CallWindowProcA(m_VMemPtr, sum1, sum2, 0, 0)
    
End Sub
