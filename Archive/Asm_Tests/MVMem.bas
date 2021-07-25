Attribute VB_Name = "MVMem"
Option Explicit

Private Declare Function VirtualAlloc Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
                         
Private Declare Function VirtualFree Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
                         
Private Declare Function VirtualLock Lib "kernel32.dll" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long) As Long
                         
Private Const MEM_COMMIT As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const MEM_RELEASE As Long = &H8000&

Public Declare Sub CopyMemory Lib "kernel32.dll" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private m_VMemPtr As Long
Private m_MemSize As Long

Public Function VMemPtr(ByVal aSize As Long) As Long
    m_MemSize = aSize
    VMemPtr = VirtualAlloc(ByVal 0, aSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    m_VMemPtr = VMemPtr
    VirtualLock ByVal VMemPtr, aSize
End Function

Public Sub VMemPtrFree(aVMemPtr As Long)
    VirtualFree ByVal aVMemPtr, 0, MEM_RELEASE
End Sub
