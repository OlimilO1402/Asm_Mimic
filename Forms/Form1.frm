VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   7335
   End
   Begin VB.CommandButton BtnTestPop1 
      Caption         =   "Test pop"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestPush1 
      Caption         =   "Test push"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestAdd1 
      Caption         =   "Test add"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestMov1 
      Caption         =   "Test mov"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestPop0 
      Caption         =   "Test pop"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestPush0 
      Caption         =   "Test push"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestAdd0 
      Caption         =   "Test add"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestMov0 
      Caption         =   "Test mov"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Realtime-Assembler MAsm:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Mockup mimicing Asm:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
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
    
    List1.Clear
    
    'mov(dst_mem_reg, src_imm_mem_reg)
    
    With Asm
        
        List1.AddItem "; mov dst_mem_reg, src_imm_mem_reg"
        
        Dim val1 As Long
        List1.AddItem "Dim val1 As Long: val1 = " & val1
        
        'mov  mem, imm
        .mov val1, 24
        List1.AddItem "mov val1, 24     ; => val1 = " & val1
        
        Dim val2 As Long: val2 = 33
        List1.AddItem "Dim val2 As Long: val2 = " & val2
       
        'mov  mem, mem
        .mov val1, val2
        List1.AddItem "mov val1, val2   ; => val1 = " & val1
        
        'mov reg, imm
        .mov eax, 24
        List1.AddItem "mov eax, 24      ; => CPU.Register(eax): " & CPU.Register(eax)
       
        'mov  mem, reg
        .mov val1, eax
        List1.AddItem "mov val1, val2   ; => val1 = " & val1
        
        'mov reg, mem
        .mov eax, val2
        List1.AddItem "mov eax, val2    ; => CPU.Register(eax): " & CPU.Register(eax)
        
        'reg 8-Bit lo
        .mov al, val1
        List1.AddItem "mov al, val1     ; => Hex(CPU.Register(al)):  &H" & Hex(CPU.Register(al))
        
        'reg 8-Bit hi
        .mov ah, val2
        List1.AddItem "mov ah, val2     ; => Hex(CPU.Register(ah)):  &H" & Hex(CPU.Register(ah))
        
        'reg 16-Bit
        .mov bx, ax
        List1.AddItem "mov bx, ax       ; => Hex(CPU.Register(bx)):  &H" & Hex(CPU.Register(bx))
        
        'reg 32-Bit
        .mov eax, ebx
        List1.AddItem "mov eax, ebx     ; => Hex(CPU.Register(eax)): &H" & Hex(CPU.Register(eax))
        
    End With
'einfach ein asm-proggy in den hex-editor laden und nach dem ersten nicht-null-byte suchen
'
'mov eax, 0x00000001  ; B8 01 00 00 00
'mov ecx, 0x00000002  ; B9 02 00 00 00
'mov edx, 0x00000003  ; BA 03 00 00 00
'mov ebx, 0x00000004  ; BB 04 00 00 00
'mov ax, 0x0001       ; 66 B8 01 00
'mov cx, 0x0002       ; 66 B9 02 00
'mov dx, 0x0003       ; 66 BA 03 00
'mov bx, 0x0004       ; 66 BB 04 00
'mov ah, 0x01         ; B4 01
'mov ch, 0x02         ; B5 02
'mov dh, 0x03         ; B6 03
'mov bh, 0x04         ; B7 04
'mov al, 0x01         ; B0 01
'mov cl, 0x02         ; B1 02
'mov dl, 0x03         ; B2 03
'mov bl, 0x04         ; B3 04
    
End Sub

Private Sub BtnTestAdd0_Click()
    
    List1.Clear
    
    'add(dst_mem_reg, src_imm_mem_reg)
    
    With Asm
        
        List1.AddItem "; add dst_mem_reg, src_imm_mem_reg"
        
        Dim val1 As Long: val1 = 42
        List1.AddItem "Dim val1 As Long: val1 = " & val1
        
        'add  mem, imm
        .Add val1, 35
        List1.AddItem "add val1, 35     ; => val1 = " & val1
    
        Dim val2 As Long: val2 = 23
        List1.AddItem "Dim val2 As Long: val2 = " & val2
        
        'add  mem, mem
        .Add val1, val2
        List1.AddItem "add val1, val2   ; => val1 = " & val1
        
        .mov eax, 24
        List1.AddItem "mov eax, 24      ; => CPU.Register(eax): " & CPU.Register(eax)
        
        'add  mem, reg
        .Add val1, eax
        List1.AddItem "add val1, eax    ; => val1 = " & val1
        
        'add reg, imm
        .Add eax, 33
        List1.AddItem "add eax, 33      ; => CPU.Register(eax): " & CPU.Register(eax)
        
        .mov val1, 56
        List1.AddItem "mov val1, 56     ; => val1 = " & val1
        
        'add reg, mem
        .Add eax, val2
        List1.AddItem "add eax, val2    ; => CPU.Register(eax): " & CPU.Register(eax)
        
        'add reg, imm
        .mov ecx, eax
        List1.AddItem "mov ecx, eax     ; => CPU.Register(ecx): " & CPU.Register(ecx)
        
        'add reg, reg
        .Add eax, ecx
        List1.AddItem "add eax, ecx     ; => CPU.Register(eax): " & CPU.Register(eax)
        
       'some fancy
       'add ebx, [edx + 2 * ecx + 32]
       
       'Add ebx, (edx + 2 * ecx + 32) ' ???
        
    End With
    
End Sub

Private Sub BtnTestPush0_Click()

    List1.Clear
    
    'Push(src_imm_mem_reg)
    
    With Asm
        
        List1.AddItem "; push src_imm_mem_reg"
        
        'push imm
        .Push 142
        List1.AddItem "push 142         ; Stack.Peek: " & Stack.Peek
        
        Dim val1 As Long: val1 = 42
        List1.AddItem "Dim val1 As Long: val1 = " & val1
        
        'push mem
        .Push val1
        List1.AddItem "push val1        ; Stack.Peek: " & Stack.Peek
        
        mov eax, val1
        List1.AddItem "mov eax, val1    ; CPU.Register(eax): " & CPU.Register(eax)
        
        'push reg
        .Push eax
        List1.AddItem "push eax         ; Stack.Peek: " & Stack.Peek
                    
    End With
    
End Sub

Private Sub BtnTestPop0_Click()

    List1.Clear
    
    'Pop (dst_mem_reg)
    
    With Asm
        
        Dim val As Long
        
        .Pop val
        List1.AddItem "pop val         ; val = " & val
        
        .Pop eax
        List1.AddItem "pop eax         ; CPU.Register(eax): " & CPU.Register(eax)
        
    End With
    
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

