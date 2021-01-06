VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public eax As r32
Public ecx As r32
Public edx As r32
Public ebx As r32
Public esp As r32
Public ebp As r32
Public esi As r32
Public edi As r32

Private Sub Class_Initialize()
    Set eax = MNew.r32(er32.eax_)
    Set ecx = MNew.r32(er32.ecx_)
    Set edx = MNew.r32(er32.edx_)
    Set ebx = MNew.r32(er32.ebx_)
    Set esp = MNew.r32(er32.esp_)
    Set ebp = MNew.r32(er32.ebp_)
    Set esi = MNew.r32(er32.esi_)
    Set edi = MNew.r32(er32.edi_)
End Sub

Public Sub Add(dst_mem_reg, src_imm_mem_reg)
'page Vol. 2A 3-31
'Adds the destination operand (first operand) and the source operand (second operand) and then
'stores the result in the destination operand.
'The destination operand can be a register or a memory location;
'The source operand can be an immediate, a register, or a memory location.
'(However, two memory operands cannot be used in one instruction.)
'When an immediate value is used as an operand, it is sign-extended to the length of the destination operand format.
'
'The ADD instruction performs integer addition.
'It evaluates the result for both signed and unsigned integer operands and sets the CF and OF flags to indicate a carry (overflow) in the signed or unsigned result, respectively.
'The SF flag indicates the sign of the signed result.
'This instruction can be used with a LOCK prefix to allow the instruction to be executed atomically.
    Dim dstReg As r32
    Dim srcReg As r32
    If TypeOf dst_mem_reg Is r32 Then
        Set dstReg = dst_mem_reg
        If TypeOf src_imm_mem_reg Is r32 Then
            Set srcReg = src_imm_mem_reg
            CPU.Register(dstReg) = CPU.Register(dstReg) + CPU.Register(srcReg)
        Else
            CPU.Register(dstReg) = CPU.Register(dstReg) + src_imm_mem_reg
        End If
    Else
        If TypeOf src_imm_mem_reg Is r32 Then
            Set srcReg = src_imm_mem_reg
            dst_mem_reg = dst_mem_reg + CPU.Register(srcReg)
        Else
            dst_mem_reg = dst_mem_reg + src_imm_mem_reg
        End If
    End If
End Sub

Public Sub mov(dst_mem_reg, src_imm_mem_reg)
'page Vol. 2B 4-35
'Copies the second operand (source operand) to the first operand (destination operand).
'The source operand can be an immediate value, general-purpose register, segment register, or memory location;
'the destination register can be a general-purpose register, segment register, or memory location.
'
'Both operands must be the same size, which can be a byte, a word, a doubleword, or a quadword.
'The MOV instruction cannot be used to load the CS register. Attempting to do so results in an invalid opcode excep-
'tion (#UD). To load the CS register, use the far JMP, CALL, or RET instruction.
'If the destination operand is a segment register (DS, ES, FS, GS, or SS), the source operand must be a valid
'segment selector. In protected mode, moving a segment selector into a segment register automatically causes the
'segment descriptor information associated with that segment selector to be loaded into the hidden (shadow) part
'of the segment register. While loading this information, the segment selector and segment descriptor information
'is validated (see the “Operation” algorithm below). The segment descriptor data is obtained from the GDT or LDT
'entry for the specified segment selector.
'A NULL segment selector (values 0000-0003) can be loaded into the DS, ES, FS, and GS registers without causing
'a protection exception. However, any subsequent attempt to reference a segment whose corresponding segment
'register is loaded with a NULL value causes a general protection exception (#GP) and no memory reference occurs.
'Loading the SS register with a MOV instruction suppresses or inhibits some debug exceptions and inhibits inter-
'rupts on the following instruction boundary. (The inhibition ends after delivery of an exception or the execution of
'the next instruction.) This behavior allows a stack pointer to be loaded into the ESP register with the next instruc-
'tion (MOV ESP, stack-pointer value) before an event can be delivered. See Section 6.8.3, “Masking Exceptions
'and Interrupts When Switching Stacks,” in Intel® 64 and IA-32 Architectures Software Developer’s Manual,
'Volume 3A. Intel recommends that software use the LSS instruction to load the SS register and ESP together.
    Dim dstReg As r32
    Dim srcReg As r32
    If TypeOf dst_mem_reg Is r32 Then
        Set dstReg = dst_mem_reg
        If TypeOf src_imm_mem_reg Is r32 Then
            Set srcReg = src_imm_mem_reg
            CPU.Register(dstReg) = MComputer.CPU.Register(srcReg)
        Else
            CPU.Register(dstReg) = src_imm_mem_reg
        End If
    Else
        If TypeOf src_imm_mem_reg Is r32 Then
            Set srcReg = src_imm_mem_reg
            dst_mem_reg = MComputer.CPU.Register(srcReg)
        Else
            dst_mem_reg = src_imm_mem_reg
        End If
    End If

End Sub

Public Sub Pop(dst_mem_reg)
'page 4-390 Vol. 2B
'Loads the value from the top of the stack to the location specified with the destination operand (or explicit opcode)
'and then increments the stack pointer. The destination operand can be a general-purpose register, memory loca-
'tion, or segment register.
'Address and operand sizes are determined and used as follows:
'• Address size. The D flag in the current code-segment descriptor determines the default address size; it may be
'overridden by an instruction prefix (67H).
'The address size is used only when writing to a destination operand in memory.
'• Operand size. The D flag in the current code-segment descriptor determines the default operand size; it may
'be overridden by instruction prefixes (66H or REX.W).
'The operand size (16, 32, or 64 bits) determines the amount by which the stack pointer is incremented (2, 4
'or 8).
'• Stack-address size. Outside of 64-bit mode, the B flag in the current stack-segment descriptor determines the
'size of the stack pointer (16 or 32 bits); in 64-bit mode, the size of the stack pointer is always 64 bits.
'The stack-address size determines the width of the stack pointer when reading from the stack in memory and
'when incrementing the stack pointer. (As stated above, the amount by which the stack pointer is incremented
'is determined by the operand size.)
'If the destination operand is one of the segment registers DS, ES, FS, GS, or SS, the value loaded into the register
'must be a valid segment selector. In protected mode, popping a segment selector into a segment register automat-
'ically causes the descriptor information associated with that segment selector to be loaded into the hidden
'(shadow) part of the segment register and causes the selector and the descriptor information to be validated (see
'the “Operation” section below).
'A NULL value (0000-0003) may be popped into the DS, ES, FS, or GS register without causing a general protection
'fault. However, any subsequent attempt to reference a segment whose corresponding segment register is loaded
'with a NULL value causes a general protection exception (#GP). In this situation, no memory reference occurs and
'the saved value of the segment register is NULL.
'The POP instruction cannot pop a value into the CS register. To load the CS register from the stack, use the RET
'instruction.
'If the ESP register is used as a base register for addressing a destination operand in memory, the POP instruction
'computes the effective address of the operand after it increments the ESP register. For the case of a 16-bit stack
'where ESP wraps to 0H as a result of the POP instruction, the resulting location of the memory write is processor-
'family-specific.
'The POP ESP instruction increments the stack pointer (ESP) before data at the old top of stack is written into the
'destination.
'Loading the SS register with a POP instruction suppresses or inhibits some debug exceptions and inhibits interrupts
'on the following instruction boundary. (The inhibition ends after delivery of an exception or the execution of the
'next instruction.) This behavior allows a stack pointer to be loaded into the ESP register with the next instruction
'(POP ESP) before an event can be delivered. See Section 6.8.3, “Masking Exceptions and Interrupts When
'Switching Stacks,” in Intel® 64 and IA-32 Architectures Software Developer’s Manual, Volume 3A. Intel recom-
'mends that software use the LSS instruction to load the SS register and ESP together.
'In 64-bit mode, using a REX prefix in the form of REX.R permits access to additional registers (R8-R15). When in
'64-bit mode, POPs using 32-bit operands are not encodable and POPs to DS, ES, SS are not valid. See the
'summary chart at the beginning of this section for encoding data and limits.
    Dim dstReg As r32
    If TypeOf dst_mem_reg Is r32 Then
        Set dstReg = dst_mem_reg
        CPU.Register(dstReg) = Stack.Pop
    Else
        dst_mem_reg = Stack.Pop
    End If
End Sub

Public Sub Push(src_imm_mem_reg)
'page Vol. 2B 4-513
'Decrements the stack pointer and then stores the source operand on the top of the stack. Address and operand
'sizes are determined and used as follows:
'• Address size. The D flag in the current code-segment descriptor determines the default address size; it may be
'overridden by an instruction prefix (67H).
'The address size is used only when referencing a source operand in memory.
'• Operand size. The D flag in the current code-segment descriptor determines the default operand size; it may
'be overridden by instruction prefixes (66H or REX.W).
'The operand size (16, 32, or 64 bits) determines the amount by which the stack pointer is decremented (2, 4 or 8).
'If the source operand is an immediate of size less than the operand size, a sign-extended value is pushed on
'the stack. If the source operand is a segment register (16 bits) and the operand size is 64-bits, a zero-
'extended value is pushed on the stack; if the operand size is 32-bits, either a zero-extended value is pushed
'on the stack or the segment selector is written on the stack using a 16-bit move. For the last case, all recent
'Core and Atom processors perform a 16-bit move, leaving the upper portion of the stack location unmodified.
'• Stack-address size. Outside of 64-bit mode, the B flag in the current stack-segment descriptor determines the
'size of the stack pointer (16 or 32 bits); in 64-bit mode, the size of the stack pointer is always 64 bits.
'The stack-address size determines the width of the stack pointer when writing to the stack in memory and
'when decrementing the stack pointer. (As stated above, the amount by which the stack pointer is
'decremented is determined by the operand size.)
'If the operand size is less than the stack-address size, the PUSH instruction may result in a misaligned stack
'pointer (a stack pointer that is not aligned on a doubleword or quadword boundary).
'The PUSH ESP instruction pushes the value of the ESP register as it existed before the instruction was executed. If
'a PUSH instruction uses a memory operand in which the ESP register is used for computing the operand address,
'the address of the operand is computed before the ESP register is decremented.
'If the ESP or SP register is 1 when the PUSH instruction is executed in real-address mode, a stack-fault exception
'(#SS) is generated (because the limit of the stack segment is violated). Its delivery encounters a second stack-
'fault exception (for the same reason), causing generation of a double-fault exception (#DF). Delivery of the
'double-fault exception encounters a third stack-fault exception, and the logical processor enters shutdown mode.
'See the discussion of the double-fault exception in Chapter 6 of the Intel® 64 and IA-32 Architectures Software
'Developer’s Manual, Volume 3A.
    Dim srcReg As r32
    If TypeOf src_imm_mem_reg Is r32 Then
        Set srcReg = src_imm_mem_reg
        Stack.Push CPU.Register(srcReg)
    Else
        Stack.Push src_imm_mem_reg
    End If
End Sub


