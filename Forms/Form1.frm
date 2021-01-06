VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestPop1 
      Caption         =   "Test pop"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestPush1 
      Caption         =   "Test push"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestAdd1 
      Caption         =   "Test add"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestMov1 
      Caption         =   "Test mov"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestPop0 
      Caption         =   "Test pop"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestPush0 
      Caption         =   "Test push"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestAdd0 
      Caption         =   "Test add"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton BtnTestMov0 
      Caption         =   "Test mov"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Realtime-Assembler MAsm:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Mockup mimicing Asm:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Asm As Asm 'Assembler Mockup mimics an assembler in VB

Private Sub Form_Load()
    Set Asm = New Asm
End Sub

Private Sub BtnTestMov0_Click()

    'mov(dst_mem_reg, src_imm_mem_reg)
    With Asm
        
        Dim val1 As Long
        MsgBox "val1 = " & val1
        
        'mov  mem, imm
        .mov val1, 24
        MsgBox "mov val1, 24     ; => val1= " & val1
        
        Dim val2 As Long: val2 = 33
        MsgBox "val2 = " & val2
       
        'mov  mem, mem
        .mov val1, val2
        MsgBox "mov val1, val2   ; => val1= " & val1
        
        'mov reg, imm
        .mov eax, 24
        MsgBox "mov eax, 24      ; => CPU-Register EAX: " & CPU.RegEAX
       
        'mov  mem, reg
        .mov val1, eax
        MsgBox "mov val1, val2   ; => val1= " & val1
            
        'mov reg, mem
        .mov eax, val2
        MsgBox "mov eax, val2    ; => CPU-Register EAX: " & CPU.RegEAX
        
    End With
    
End Sub

Private Sub BtnTestAdd0_Click()
    
    'add(dst_mem_reg, src_imm_mem_reg)
    
    With Asm
        
        Dim val1 As Long: val1 = 42
        MsgBox "val1 = " & val1
        
       'add  mem, imm
        .Add val1, 35
        MsgBox "add val1, 35     ; => val1 = " & val1
    
        Dim val2 As Long: val2 = 23
        MsgBox "val2 = " & val2
        
       'add  mem, mem
        .Add val1, val2
        MsgBox "add val1, val2   ; => val1 = " & val1
        
        .mov eax, 24
        MsgBox "mov eax, 24      ; => CPU-Register EAX: " & CPU.RegEAX
        
       'add  mem, reg
        .Add val1, eax
        MsgBox "add val1, eax    ; => val1 = " & val1
        
       'add reg, imm
        .Add eax, 33
        MsgBox "add eax, 33      ; => CPU-Register EAX: " & CPU.RegEAX
        
        .mov val1, 56
        MsgBox "mov val1, 56     ; => val1 = " & val1
        
       'add reg, mem
        .Add eax, val2
        MsgBox "add eax, val2    ; => CPU-Register EAX: " & MComputer.CPU.RegEAX
        
        .mov ecx, 12
        MsgBox "mov ecx, 75      ; => CPU-Register ECX: " & MComputer.CPU.RegECX
        
       'add reg, reg
        .Add eax, ecx
        MsgBox "add eax, ecx     ; => CPU-Register EAX: " & MComputer.CPU.RegEAX
        
    End With
End Sub

Private Sub BtnTestPop0_Click()
    
    'Pop (dst_mem_reg)
    
End Sub

Private Sub BtnTestPush0_Click()
    
    'Push(src_imm_mem_reg)
    
End Sub



Private Sub BtnTestAdd1_Click()
    MsgBox "TODO: mnemonic -> create instruction with opcode,modR/M-Byte,SIB-Byte,etc."
End Sub
Private Sub BtnTestMov1_Click()
    MsgBox "TODO: mnemonic -> create instruction with opcode,modR/M-Byte,SIB-Byte,etc."
End Sub
Private Sub BtnTestPop1_Click()
    MsgBox "TODO: mnemonic -> create instruction with opcode,modR/M-Byte,SIB-Byte,etc."
End Sub
Private Sub BtnTestPush1_Click()
    MsgBox "TODO: mnemonic -> create instruction with opcode,modR/M-Byte,SIB-Byte,etc."
End Sub

