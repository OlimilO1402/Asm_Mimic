Attribute VB_Name = "MInstruction"
Option Explicit
'youtube e.g. Smruti R. Sarangi

Public Type OptBytes4 'Prefix, Displacement, Immediate
    Count As Byte
    Value1 As Byte
    Value2 As Byte
    Value3 As Byte
    Value4 As Byte
End Type

Public Type OptByte1 'ModR/M, SIB
    Count As Byte
    Value As Byte
End Type

Public Type Instruction
    Prefix As OptBytes4
    Opcode As OptBytes4 'the only one that is *not* optional!
    ModRM  As OptByte1  'mod-Reg-R/M-Byte
    SIB    As OptByte1  'Scale Index Base
    Displ  As OptBytes4 'Displacement Bytes 'Displacement = Offset
    Immed  As OptBytes4 'Immediate Bytes
End Type
'1 - 17 Bytes possible

Public Function New_OptByte1(ByVal aValue As Byte) As OptByte1
    With New_OptByte1
        .Count = 1
        .Value = aValue
    End With
End Function

Public Function New_ModRM(mode As Byte, bReg As Byte, bRM As Byte) As OptByte1
    With New_ModRM
        .Count = 1
        .Value = bRM + bReg * 2 ^ 3 + mode * 2 ^ 6
    End With
End Function
'| mod |  Reg  |  R/M  | (R=Register, M=Memory)
'| 8 7 | 6 5 4 | 3 2 1 |
'   2      3       3    bits
'        Oper2   Oper1                    e.g. add Oper2, Oper1 '???
'
'mod-bits: determine the adressing modes of the operands
'00 -> register indirect adressing mode (no displacement)
'      e.g. add eax, [ebx] oder add [ebx], eax : these two have the same modR/M-Byte, but opcode has different direction field 0/1
'01 -> Indirect adressing mode with 1 byte displacement
'10 -> Indirect adressing mode with 4 byte displacement
'11 -> Register direct adressing mode, both operands are registers
'      e.g. add eax, ebx
'
'00, 01, 10 -> only one operand can be a memory address
'
'if the R/M-bits are 100 (=esp) then we use the SIB-Byte
'If mode = 00, and R/M = 101 (=ebp), we use memory direct adressing the 32 bit displacement is used as the memory address
'
'Code
'000 = 0 = eax
'001 = 1 = ecx
'010 = 2 = edx
'011 = 3 = ebx
'100 = 4 = esp
'101 = 5 = ebp
'110 = 6 = esi
'111 = 7 = edi
'
'e.g. mov eax, [0xABCDEF12]


Public Function New_SIB(sc As Byte, ix As Byte, bs As Byte) As OptByte1
    With New_SIB
        .Count = 1
    End With
End Function
'| Scale | Index | Base  |
'|  8 7  | 6 5 4 | 3 2 1 |
'   2      3       3    bits
'there are 4 values of the scale: 00(1), 01(2), 10(4), 11(8)
'00: e.g. eax + ebx
'01: e.g. eax + 2 * ebx
'10: e.g. eax + 4 * ebx
'11: e.g. eax + 8 * ebx
'
'The Index and Base are 3 bits each an follow the register encoding scheme
'some rules:
'   * esp can not be an index
'   * the offset of the memory adress can only be specified in the displacement field
'
' add ebx, [edx + 2 * ecx + 32] ; assume that the opcode for the add instruction is 03
'Let us calculate   the value of the modR/M byte. In this case, our displacement fits
'within 8 bits. Hence, we can set the mod-bits equal to 01 (corresponding to an 8bit displacement)
'We need to use the SIB byte because we have a scale and an index. Thus we set the R/M-bits to 100
'The destination register is ebx. it's code is 011. Thus the ModR/M-Byte is 01 011 100 (=&H5C)
'
'Now let's calculatae the value of the SIB-byte. The scale is equal to two (01) the index is ecx(=001)
'and the base is edx(=010), Hence the SIB-byte is 01 001 010 (=&H4A)
'The last byte is the displacement, which is equal to 0x20 '
'Thus the encoding of the instruction is 03 5C 4A 20 (in hex)



'
Public Function New_Instruction(pfx As OptBytes4, opc As OptBytes4, mrm As OptByte1, si As OptByte1, dis As OptBytes4, imm As OptBytes4) As Instruction
    With New_Instruction: .Prefix = pfx: .Opcode = opc:    .ModRM = mrm:      .SIB = si:    .Displ = dis:     .Immed = imm:    End With
End Function

Public Function Instruction_Count(Instruc As Instruction) As Byte
    Dim c As Byte
    With Instruc
        c = .Prefix.Count + .Opcode.Count + .ModRM.Count + .SIB.Count + .Displ.Count + .Immed.Count
    End With
    Instruction_Count = c
End Function
