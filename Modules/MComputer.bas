Attribute VB_Name = "MComputer"
Option Explicit
Public CPU   As CPU
Public Stack As Stack
Public Heap  As Memory

Public Sub Init()
    Set CPU = New CPU
    Set Stack = New Stack
    Set Heap = New Memory
End Sub
