;Dither.asm  by Robert Rayment  Nov 2002

; 26/11/02  Using a Jump Table for Subs

; ASM layout done to roughly match VB layout

; VB
; Public ptrMC, ptrStruc    ' Ptrs to Machine Code & Structure
; 'MCode Structure
; Public Type PStruc
;   picWd As Long
;   picHt As Long
;   Ptrpic1mem As Long      ' pic1mem(1-4,1-picWd,1-picHt) bytes
;   Ptrpic2mem As Long      ' pic2mem(1-4,1-picWd,1-picHt) bytes
;   PtrIArr As Long         ' IArr(1-picWd,1-picHt)
;   BlackThreshHold As Long ' 0-255
;   DithDIV As Long         ' FSADIV, FSBDIV,StuckDiv etc
;   greysum As Long
; '---------------------------------
;   NX As Long
;   NY As Long
;   NT As Long
;   PtrLimits As Long       ' Limits(1-NT)
;   PtrTM As Long           ' TM(1-NX, 1-NY, 1-NT) bytes
;   RandSeed As Long        ' (0-.99999) * 255
;   OpCode As Long
; End Type
; Public DStruc As PStruc
; EG
; Public DitherMC() As Byte
; Public Sub ASM_BlackWhite()
; DStruc.OpCode = 0
; res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
;                             8         12  16  20
; End Sub

%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%define picWd          [ebp-4]
%define picHt          [ebp-8]
%define Ptrpic1mem     [ebp-12]
%define Ptrpic2mem     [ebp-16]
%define PtrIArr        [ebp-20]
%define BlackThreshHold [ebp-24]
%define DithDIV        [ebp-28]
%define greysum        [ebp-32] ; Also AvGrey
%define NX             [ebp-36]
%define NY             [ebp-40]
%define NT             [ebp-44]
%define PtrLimits      [ebp-48]
%define PtrTM          [ebp-52]
%define RandSeed       [ebp-56]
%define OpCode         [ebp-60]
; Some variables
%define j       [ebp-64]
%define i       [ebp-68]

%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]
%define N       [ebp-88]
%define jj      [ebp-92]
%define ii      [ebp-96]
%define jjup    [ebp-100]
%define iiup    [ebp-104]
%define cul0    [ebp-108]
%define culr    [ebp-112]
%define Seed    [ebp-116] ; from RandSeed
%define Rand    [ebp-120]
%define T255    [ebp-124]

%define ptmc    [ebp-128]   ; ptr mcode buffer



[bits 32]

    push ebp
    mov ebp,esp
    sub esp,128
    push edi
    push esi
    push ebx
	
	; Get absolute address to mcode buffer in VB
    Call BufferADDR
    sub eax,17          ; 17B from push ebp to CallBufferADDR incl
    mov ptmc,eax


    ;Fill structure
    mov ebx,            [ebp+8]
    movab picWd,        [ebx]
    movab picHt,        [ebx+4]
    movab Ptrpic1mem,   [ebx+8]
    movab Ptrpic2mem,   [ebx+12]
    movab PtrIArr,      [ebx+16]
    movab BlackThreshHold,[ebx+20]
    movab DithDIV,      [ebx+24]
    movab greysum,      [ebx+28]
    movab NX,           [ebx+32]
    movab NY,           [ebx+36]
    movab NT,           [ebx+40]
    movab PtrLimits,    [ebx+44]
    movab PtrTM,        [ebx+48]
    movab RandSeed,     [ebx+52]
    movab OpCode,       [ebx+56]
;----------------------------

	; Using jump table
    mov ebx,Sub0
    mov eax,OpCode
	cmp eax,8
	jg GETOUT			; Error
    shl eax,2           ; x4
    add ebx,eax
    add ebx,ptmc
    mov eax,[ebx]
    add eax,ptmc
    Call eax

;    jmp GETOUT
;    
;
;    mov eax,OpCode
;    cmp eax,0
;    jne T1
;    Call BLACKWHITE
;    jmp GETOUT
;T1:
;    cmp eax,1
;    jne T2
;    Call FloydSteinbergA
;    jmp GETOUT
;
;T2: 
;    cmp eax,2
;    jne T3
;    Call FloydSteinbergB
;    jmp GETOUT
;T3: 
;    cmp eax,3
;    jne T4
;    Call Stucki
;    jmp GETOUT
;T4:
;    cmp eax,4
;    jne T5
;    Call Burkes
;    jmp GETOUT
;T5:
;    cmp eax,5
;    jne T6
;    Call Sierra
;    jmp GETOUT
;T6:
;    cmp eax,6
;    jne T7
;    Call Jarvis
;    jmp GETOUT
;T7:
;    cmp eax,7
;    jne T8
;    Call DOPATTERN
;    jmp GETOUT
;T8:
;    cmp eax,8
;    jne GETOUT
;    Call RANDOM
;
;    jmp GETOUT


GETOUT:
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

;=====================================================================
; Get mcode buffer absolute start addres
BufferADDR:
    pop eax
    push eax
RET
;=====================================================

BLACKWHITE:
    ; Zero pic2mem
    mov eax,picHt
    mov ebx,picWd
    mul ebx
    mov ecx,eax         ; picHt*picWd no. of 4-bytes chunks in pic2mem
    xor eax,eax         ; eax=0
    mov edi,Ptrpic2mem
    rep stosd
 
    mov ecx,picHt
ForjBW:
    mov j,ecx
    push ecx
    mov ecx,picWd
ForiBW:
    mov i,ecx
    mov edi,Ptrpic1mem
    Call GetAddrpicji       ; edi-> pic1mem(1,i,j)
    xor eax,eax
    xor ebx,ebx
    mov AL,Byte[edi]    ; B
    mov BL,Byte[edi+1]  ; G
    add eax,ebx
    mov BL,Byte[edi+2]  ; R
    add eax,ebx
    mov ebx,3
    div ebx             ; eax = (B+G+R)\3
    cmp eax,BlackThreshHold
    jle NextiBW         ; AL<=BlackThreshHold 

    mov edi,Ptrpic2mem
    Call GetAddrpicji       ; edi-> pic2mem(1,i,j)
    mov [edi],DWord 0FFFFFFh    ; White
NextiBW:
    dec ecx
    jnz ForiBW

    pop ecx
    dec ecx
    jnz ForjBW
RET

;=====================================================
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

FloydSteinbergA:

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; FSADIV,1
    fdivp st1           ; st1/sto = 1/FSADIV
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
FSAj:
    mov j,ecx
    push ecx
    mov ecx,1
FSAi:
    mov i,ecx
    
    Call Fillpic2memGetzErr
    
    ; Spread error j=j+1
    inc DWord j         ; j=j+1

    mov eax,3
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+3*zErr

    inc DWord i         ; i=i+1
    mov eax,2
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+2*zErr
    dec DWord i         ; i=i

    dec DWord j         ; j=j

    ; Spread error i=i+1
    inc DWord i         ; i=i+1

    mov eax,3
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+3*zErr

    dec DWord i         ; i=i  ; Some of these could be omitted

NextFSAi:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near FSAi

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near FSAj
RET

;=====================================================
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

FloydSteinbergB:

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; FSBDIV,1
    fdivp st1           ; st1/sto = 1/FSBDIV
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
FSBj:
    mov j,ecx
    push ecx
    mov ecx,1
FSBi:
    mov i,ecx

    Call Fillpic2memGetzErr

    ; Spread error j=j+1
    inc DWord j         ; j=j+1 ===

    dec DWord i         ; i=i-1
    mov eax,3
    Call Diffuse        ; IArr(i-1,j+1)=IArr(i-1,j+1)+3*zErr
    
    inc DWord i         ; i=i
    mov eax,5
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+5*zErr
    
    inc DWord i         ; i=i+1
    mov eax,1
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+1*zErr
    dec DWord i         ; i=i-1

    dec DWord j         ; j=j ===

    ; Spread error j=j
    inc DWord i         ; i=i+1
    mov eax,7
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+7*zErr
    dec DWord i         ; i=i

NextFSBi:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near FSBi

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near FSBj
RET

;=====================================================
Stucki:
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; Stucki,1
    fdivp st1           ; st1/sto = 1/Stucki
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
Stuckij:
    mov j,ecx
    push ecx
    mov ecx,1

Stuckii:
    mov i,ecx

    Call Fillpic2memGetzErr

    ; Spread error j+2
    inc DWord j
    inc DWord j         ; j=j+2 ===
    
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,1
    Call Diffuse        ; IArr(i-2,j+2)=IArr(i-2,j+1)+1*zErr
    inc DWord i         ; i=i-1
    mov eax,2
    Call Diffuse        ; IArr(i-1,j+2)=IArr(i-1,j+1)+2*zErr
    inc DWord i         ; i=i
    mov eax,4
    Call Diffuse        ; IArr(i,j+2)=IArr(i,j+1)+4*zErr
    inc DWord i         ; i=i+1
    mov eax,2
    Call Diffuse        ; IArr(i+1,j+2)=IArr(i+1,j+1)+2*zErr
    inc DWord i         ; i=i+2
    mov eax,1
    Call Diffuse        ; IArr(i+2,j+2)=IArr(i+2,j+1)+2*zErr
    dec DWord i
    dec DWord i         ; i=i

    ; Spread error j+1
    dec DWord j         ; j=j+1
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,2
    Call Diffuse        ; IArr(i-2,j+1)=IArr(i-2,j+1)+2*zErr
    inc DWord i         ; i=i-1
    mov eax,4
    Call Diffuse        ; IArr(i-1,j+1)=IArr(i-1,j+1)+4*zErr
    inc DWord i         ; i=i
    mov eax,8
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+8*zErr
    inc DWord i         ; i=i+1
    mov eax,4
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+4*zErr
    inc DWord i         ; i=i+2
    mov eax,2
    Call Diffuse        ; IArr(i+2,j+1)=IArr(i+2,j+1)+2*zErr
    dec DWord i         ; i=i+1

    ; Spread error j
    dec DWord j         ; j=j
    mov eax,8
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+8*zErr
    inc DWord i         ; i=i+2
    mov eax,4
    Call Diffuse        ; IArr(i+2,j)=IArr(i+2,j)+4*zErr
    dec DWord i
    dec DWord i         ; i=i

NextStuckii:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near Stuckii

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near Stuckij
RET

;=====================================================
Burkes:
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; Burkes,1
    fdivp st1           ; st1/sto = 1/Burkes
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
Burkesj:
    mov j,ecx
    push ecx
    mov ecx,1

Burkesi:
    mov i,ecx

    Call Fillpic2memGetzErr

    ; Spread error j+1
    inc DWord j         ; j=j+1
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,2
    Call Diffuse        ; IArr(i-2,j+1)=IArr(i-2,j+1)+2*zErr
    inc DWord i         ; i=i-1
    mov eax,4
    Call Diffuse        ; IArr(i-1,j+1)=IArr(i-1,j+1)+4*zErr
    inc DWord i         ; i=i
    mov eax,8
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+8*zErr
    inc DWord i         ; i=i+1
    mov eax,4
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+4*zErr
    inc DWord i         ; i=i+2
    mov eax,2
    Call Diffuse        ; IArr(i+2,j+1)=IArr(i+2,j+1)+2*zErr
    dec DWord i         ; i=i+1

    ; Spread error j
    dec DWord j         ; j=j
    mov eax,8
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+8*zErr
    inc DWord i         ; i=i+2
    mov eax,4
    Call Diffuse        ; IArr(i+2,j)=IArr(i+2,j)+4*zErr
    dec DWord i
    dec DWord i         ; i=i

NextBurkesi:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near Burkesi

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near Burkesj
RET

;=====================================================
Sierra:
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; Sierra,1
    fdivp st1           ; st1/sto = 1/Sierra
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
Sierraj:
    mov j,ecx
    push ecx
    mov ecx,1

Sierrai:
    mov i,ecx

    Call Fillpic2memGetzErr

    ; Spread error j+2
    inc DWord j
    inc DWord j         ; j=j+2 ===
    
    dec DWord i         ; i=i-1
    mov eax,2
    Call Diffuse        ; IArr(i-1,j+2)=IArr(i-1,j+1)+2*zErr
    inc DWord i         ; i=i
    mov eax,3
    Call Diffuse        ; IArr(i,j+2)=IArr(i,j+1)+3*zErr
    inc DWord i         ; i=i+1
    mov eax,2
    Call Diffuse        ; IArr(i+1,j+2)=IArr(i+1,j+1)+2*zErr
    dec DWord i         ; i=i

    ; Spread error j+1
    dec DWord j         ; j=j+1
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,2
    Call Diffuse        ; IArr(i-2,j+1)=IArr(i-2,j+1)+2*zErr
    inc DWord i         ; i=i-1
    mov eax,4
    Call Diffuse        ; IArr(i-1,j+1)=IArr(i-1,j+1)+4*zErr
    inc DWord i         ; i=i
    mov eax,5
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+5*zErr
    inc DWord i         ; i=i+1
    mov eax,4
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+4*zErr
    inc DWord i         ; i=i+2
    mov eax,2
    Call Diffuse        ; IArr(i+2,j+1)=IArr(i+2,j+1)+2*zErr
    dec DWord i         ; i=i+1

    ; Spread error j
    dec DWord j         ; j=j
    mov eax,5
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+5*zErr
    inc DWord i         ; i=i+2
    mov eax,3
    Call Diffuse        ; IArr(i+2,j)=IArr(i+2,j)+3*zErr
    dec DWord i
    dec DWord i         ; i=i

NextSierrai:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near Sierrai

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near Sierraj
RET

;=====================================================
Jarvis:
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

    Call FillIntensityArray

    fld1
    fild DWord DithDIV  ; Jarvis,1
    fdivp st1           ; st1/sto = 1/Jarvis
    fstp DWord zMul

    mov ecx,1           ; Need to start at 1
Jarvisj:
    mov j,ecx
    push ecx
    mov ecx,1

Jarvisi:
    mov i,ecx

    Call Fillpic2memGetzErr

    ; Spread error j+2
    inc DWord j
    inc DWord j         ; j=j+2 ===
    
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,1
    Call Diffuse        ; IArr(i-2,j+2)=IArr(i-2,j+1)+1*zErr
    inc DWord i         ; i=i-1
    mov eax,3
    Call Diffuse        ; IArr(i-1,j+2)=IArr(i-1,j+1)+3*zErr
    inc DWord i         ; i=i
    mov eax,5
    Call Diffuse        ; IArr(i,j+2)=IArr(i,j+1)+5*zErr
    inc DWord i         ; i=i+1
    mov eax,3
    Call Diffuse        ; IArr(i+1,j+2)=IArr(i+1,j+1)+3*zErr
    inc DWord i         ; i=i+2
    mov eax,1
    Call Diffuse        ; IArr(i+2,j+2)=IArr(i+2,j+1)+2*zErr
    dec DWord i
    dec DWord i         ; i=i

    ; Spread error j+1
    dec DWord j         ; j=j+1
    dec DWord i
    dec DWord i         ; i=i-2
    mov eax,3
    Call Diffuse        ; IArr(i-2,j+1)=IArr(i-2,j+1)+3*zErr
    inc DWord i         ; i=i-1
    mov eax,5
    Call Diffuse        ; IArr(i-1,j+1)=IArr(i-1,j+1)+5*zErr
    inc DWord i         ; i=i
    mov eax,7
    Call Diffuse        ; IArr(i,j+1)=IArr(i,j+1)+7*zErr
    inc DWord i         ; i=i+1
    mov eax,5
    Call Diffuse        ; IArr(i+1,j+1)=IArr(i+1,j+1)+5*zErr
    inc DWord i         ; i=i+2
    mov eax,3
    Call Diffuse        ; IArr(i+2,j+1)=IArr(i+2,j+1)+2*zErr
    dec DWord i         ; i=i+1

    ; Spread error j
    dec DWord j         ; j=j
    mov eax,7
    Call Diffuse        ; IArr(i+1,j)=IArr(i+1,j)+8*zErr
    inc DWord i         ; i=i+2
    mov eax,5
    Call Diffuse        ; IArr(i+2,j)=IArr(i+2,j)+4*zErr
    dec DWord i
    dec DWord i         ; i=i

NextJarvisi:
    inc ecx             ; i+1
    cmp ecx,picWd
    jle Near Jarvisi

    pop ecx
    inc ecx             ; j+1
    cmp ecx,picHt
    jle Near Jarvisj
RET

;=====================================================
%define zMul    [ebp-72]    
%define cul     [ebp-76]
%define zErr    [ebp-80]
%define Temp    [ebp-84]

Diffuse:

    mov Temp,eax
    fld DWord zErr
    fild DWord Temp
    fmulp st1           ; #*zErr
    mov edi,PtrIArr
    Call GetAddrIArrji  ; edi-> IArr(i*,j*)
    fild DWord [edi]
    faddp st1
    fistp DWord [edi]   ; IArr(i*,j*)=IArr(i*,j*)+#*zErr

RET

;=====================================================

%define j       [ebp-64]
%define i       [ebp-68]
%define jup     [ebp-72]    
%define iup     [ebp-76]
%define AvCount [ebp-80]
%define AvGrey  [ebp-84]
%define N       [ebp-88]
%define jj      [ebp-92]
%define ii      [ebp-96]

DOPATTERN:
;   NX,NY,NT,PtrLimits

    Call FillIntensityArray
    mov eax,1
ForJJ:
    mov jj,eax
    add eax,NY
    dec eax
    mov jup,eax

    mov eax,1
ForII:
    mov ii,eax
    add eax,NX
    dec eax
    mov iup,eax

    Call GetAvGrey
    mov eax,AvCount
    je Near NexJJ

    mov eax,1
ForN:
    mov N,eax

    mov edi,PtrLimits
    dec eax
    shl eax,2
    add edi,eax
    mov eax,[edi]
    cmp eax,AvGrey
    jge NexN

    mov ecx,jj
Forj2:
    cmp ecx,picHt
    jg NexN
    mov j,ecx
    push ecx
    mov ecx,ii
Fori2:
    cmp ecx,picWd
    jg Nexj2
    mov i,ecx
    push ecx        ; ecx wanted for GetAddrTM
        ;----------
        mov edi,PtrTM   
        Call GetAddrTM
        pop ecx
        xor eax,eax
        mov AL,Byte[edi]
        cmp AL,49
        jne Nexi2
        mov edi,Ptrpic2mem
        Call GetAddrpicji
        mov [edi],DWord 0FFFFFFh    ; White
        ;----------
Nexi2:
    inc ecx
    cmp ecx,iup
    jle Fori2
Nexj2:
    pop ecx
    inc ecx
    cmp ecx,jup
    jle Forj2

NexN:
    mov eax,N
    inc eax
    cmp eax,NT
    jle Near ForN

NexII:
    mov eax,ii
    add eax,NX
    cmp eax,picWd
    jle Near ForII
NexJJ:
    mov eax,jj
    add eax,NY
    cmp eax,picHt
    jle Near ForJJ
RET

;=====================================================
GetAvGrey:

    xor eax,eax
    mov AvCount,eax
    mov AvGrey,eax

    mov ecx,jj
Forj1:
    mov j,ecx
    push ecx
    mov ecx,ii
Fori1:
    mov i,ecx

    mov edi,PtrIArr
    Call GetAddrIArrji      ; edi->IArr(i,j)
    mov eax,AvGrey
    add eax,DWord[edi]
    mov AvGrey,eax
    inc DWord AvCount
Nexi1:
    inc ecx
    cmp ecx,iup
    jle Fori1
Nexj1:
    pop ecx
    inc ecx
    cmp ecx,jup
    jle Forj1
;------------
    mov eax,AvCount
    mov ebx,eax
    cmp eax,0
    je garet
    mov eax,AvGrey
    div ebx
    mov AvGrey,eax
garet:
RET

;=====================================================
RANDOM:
%define j       [ebp-64]
%define i       [ebp-68]
%define jup     [ebp-72]    
%define iup     [ebp-76]
%define AvCount [ebp-80]
%define AvGrey  [ebp-84]
%define N       [ebp-88]
%define jj      [ebp-92]
%define ii      [ebp-96]

%define jjup    [ebp-100]
%define iiup    [ebp-104]
%define cul0    [ebp-108]
%define culr    [ebp-112]
%define Seed    [ebp-116] ; from RandSeed
%define Rand    [ebp-120]
%define T255    [ebp-124]




    Call FillIntensityArray
    
    mov eax,255
    mov T255,eax    ; 255 for Random no.
    
    mov eax,1
ForJJR:
    mov jj,eax
    add eax,NY
    dec eax
    mov jup,eax

    mov eax,1
ForIIR:
    mov ii,eax
    add eax,NX
    dec eax
    mov iup,eax

    Call GetAvGrey
    mov eax,AvCount
    cmp eax,0
    je Near NexJJR


    mov eax,1
ForNR:
    mov N,eax

    mov edi,PtrLimits
    dec eax
    shl eax,2
    add edi,eax
    mov eax,[edi]
    cmp eax,AvGrey
    jl NexNR

    Call RandEval

    jmp TestN10

NexNR:
    mov eax,N
    inc eax
    cmp eax,NT
    jle Near ForNR

TestN10:
    mov eax,N
    cmp eax,10
    jne NexIIR
    
    Call RandEval

NexIIR:
    mov eax,ii
    add eax,NX
    cmp eax,picWd
    jle Near ForIIR
NexJJR:
    mov eax,jj
    add eax,NY
    cmp eax,picHt
    jle Near ForJJR

RET

;=====================================================
RandEval:       ; k=N, ii,jj

    xor eax,eax
    mov culr,eax
    mov eax,255
    mov cul0,eax

    mov edi,PtrLimits
    mov eax,N
    dec eax
    shl eax,2       ;(N-1)*4
    add edi,eax
    mov eax,[edi]
    cmp eax,128
    jge culset  ; RandomLimits(N)>=128
    ; Swap cul0,culr 
    xor eax,eax
    mov cul0,eax
    mov eax,255
    mov culr,eax

culset:
    mov eax,jj
    add eax,2
    mov jjup,eax

    mov eax,ii
    add eax,2
    mov iiup,eax
    

    mov eax,cul0
    cmp eax,255
    jne TCulr
    
    mov ecx,jj
ForRjj:
    mov j,ecx

    cmp ecx,picHt
    jg Ranret

    push ecx

    mov ecx,ii
ForRii:
    mov i,ecx

    cmp ecx,picWd
    jg NexRjj

        mov edi,Ptrpic2mem
        Call GetAddrpicji
        mov [edi],DWord 0FFFFFFh    ; White
NexRii:
    inc ecx
    cmp ecx,iiup
    jle ForRii

NexRjj:
    pop ecx
    inc ecx
    cmp ecx,jjup
    jle ForRjj  
RET

TCulr:

    mov eax,N
    cmp eax,0
    jle Ranret
    cmp eax,9
    jge Ranret

    mov ecx,1
Fornn:

    
    ;j=int((jjup-jj)*rnd+jj)
    ;i=int((iiup-ii)*rnd+ii)

    Call RANDNOISE

    mov eax,j
    cmp eax,picHt
    jg nexnn

    mov eax,i
    cmp eax,picWd
    jg nexnn

        mov edi,Ptrpic2mem
        Call GetAddrpicji
        mov [edi],DWord 0FFFFFFh    ; White
nexnn:
    inc ecx
    cmp ecx,N
    jl Fornn
    
Ranret:
RET

;=====================================================
FillIntensityArray:
    xor eax,eax
    mov greysum,eax     ; greysum=0

    mov ecx,picHt
FIAj:
    mov j,ecx
    push ecx
    mov ecx,picWd
FIAi:
    mov i,ecx
    mov edi,Ptrpic1mem
    Call GetAddrpicji       ; edi-> pic1mem(1,i,j)
    xor eax,eax
    xor ebx,ebx
    mov AL,Byte[edi]        ; B
    mov BL,Byte[edi+1]  ; G
    add eax,ebx
    mov BL,Byte[edi+2]  ; R
    add eax,ebx
    mov ebx,3
    div ebx             ; eax = (B+G+R)\3
    push eax                ; grey byte
    mov edi,PtrIArr     ; edi-> IArr(1,1)
    Call GetAddrIArrji      ; edi-> IArr(i,j)
    pop eax             ; grey byte
    mov [edi],eax
    add greysum,eax     ; greysum + grey byte
 
    dec ecx
    jnz FIAi

    pop ecx
    dec ecx
    jnz FIAj

    mov eax,picHt
    mov ebx,picWd
    mul ebx             ; eax=greycount
    mov ebx,eax         ; ebx=greycount
    mov eax,greysum
    div ebx             ; average greysum
    mov greysum,eax

    ; Zero pic2mem
    mov eax,picHt
    mov ebx,picWd
    mul ebx
    mov ecx,eax         ; picHt*picWd no. of 4-bytes chunks in pic2mem
    xor eax,eax         ; eax=0
    mov edi,Ptrpic2mem
    rep stosd

RET

;=====================================================
Fillpic2memGetzErr:
    xor eax,eax
    mov cul,eax         ; cul=0
    mov edi,PtrIArr
    Call GetAddrIArrji      ; edi-> IArr(i,j)
    mov eax,[edi]
    mov Temp,eax        ; IArr(i,j)
    cmp eax,greysum
    jle CalcZerr
    
    mov eax,255
    mov cul,eax         ; cul=255
    mov edi,Ptrpic2mem
    Call GetAddrpicji       ; edi-> pic2mem(1,i,j)
    mov [edi],DWord 0FFFFFFh    ; White

CalcZerr:
    fild DWord Temp     ; IArr(i,j)
    fild DWord cul
    fsubp st1           ; IArr(i,j)-cul
    fld DWord zMul      ; zMul,IArr(i,j)-cul
    fmulp st1           ; zErr
    fstp DWord zErr     ; = (IArr(i,j)-cul) * zMul
RET

;=====================================================
GetAddrIArrji:
; Zero Intensity Array
; VB ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)

;   edi-> IArr(-1,-1)  
;   addr=edi + 4*[(j+1)*picWd + (i+1)]
    mov eax,j
    inc eax

    mov ebx,picWd
    mul ebx
    mov ebx,i
    inc ebx

    add eax,ebx
    shl eax,2       ; x4
    add edi,eax
RET 

;=====================================================
GetAddrpicji:
;   edi->pic1mem(1,1,1) or pic2mem(1,1,1) 
;   addr=edi + 4*[(j-1)*picWd + (i-1)]
    mov eax,j
    dec eax

    mov ebx,picWd
    mul ebx
    mov ebx,i
    dec ebx

    add eax,ebx
    shl eax,2       ; x4
    add edi,eax
RET 

GetAddrLimits:
; edi-> Limits(1) ,N
    mov eax,N
    dec eax
    shl eax,2
    add edi,eax

;=====================================================
GetAddrTM:  ;In. edi->TM(1,1)   TM(NX,NY,N) Bytes 
            ;Out. edi->TM(i-ii+1,j-jj+1,N) byte array
            ;  NX-1  + (NY-1)*NX + (N-1)*NX*NY
            ; (i-ii) + (j-jj)*NX + (N-1)*NX*NY
    mov eax,N
    dec eax
    mov ebx,NX
    mul ebx
    mov ebx,NY
    mul ebx         ;(N-1)*NX*NY

    mov ecx,eax     ;ecx=(N-1)*NX*NY
    mov eax,j
    sub eax,jj
    mov ebx,NX
    mul ebx
    add ecx,eax     ;ecx=(j-jj)*NX + (N-1)*NX*NY

    mov eax,i
    sub eax,ii
    add eax,ecx     ;eax=(i-ii) + (j-jj)*NX + (N-1)*NX*NY
    add edi,eax
RET

;============================================================
RANDNOISE:  ;In: RandSeed,jjup,jj iiup,ii Out: j & i random postions
    mov eax,011813h     ; 71699 prime 
    imul DWORD Seed
    add eax, 0AB209h    ; 700937 prime
    rcr eax,1           ; leaving out gives vertical lines plus
                        ; faint horizontal ones, tartan

    ;----------------------------------------
    ;jc ok              ; these 2 have little effect
    ;rol eax,1          ;
ok:                     ;
    ;----------------------------------------
    
    ;----------------------------------------
    ;dec eax            ; these produce vert lines
    ;inc eax            ; & with fsin marble arches
    ;----------------------------------------

    mov Seed,eax    ; save seed
    and eax,255     ; 0-255
    mov Rand,eax
    
    fild DWord Rand    ; ran (0-255)
    fild Dword T255    ; 255
    fdivp st1           ; st1/st0  R=ran/255  (0-1)
    
    fild DWord jjup
    fild Dword jj
    fsubp st1           ; st1-st0   (jjup-jj)
    
    fmulp st1           ; (jjup-jj)*R
    fild Dword jj
    faddp st1           ; (jjup-jj)*R + jj
    fistp Dword j       ; j=Int((jjup-jj)*R + jj)

;=============================================================

    mov eax,011813h     ; 71699 prime 
    imul DWORD Seed
    add eax, 0AB209h    ; 700937 prime
    rcr eax,1           ; leaving out gives vertical lines plus
                        ; faint horizontal ones, tartan

    ;----------------------------------------
    ;jc ok2             ; these 2 have little effect
    ;rol eax,1          ;
ok2:                        ;
    ;----------------------------------------
    
    ;----------------------------------------
    ;dec eax            ; these produce vert lines
    ;inc eax            ; & with fsin marble arches
    ;----------------------------------------

    mov Seed,eax    ; save seed
    and eax,255
    mov Rand,eax
    
    fild DWord Rand    ; ran (0-255)
    fild Dword T255    ; 255
    fdivp st1           ; st1/st0   R=ran/255  (0-1)

    fild DWord iiup
    fild Dword ii
    fsubp st1           ; st1-st0   (iiup-ii)
    
    fmulp st1           ; (iiup-ii)*R
    fild Dword ii
    faddp st1           ; (iiup-ii)*R + ii
    fistp Dword i       ; i=Int((iiup-ii)*R + ii)
RET

;=============================================================
; Jump table
[SECTION .data]
Sub0 dd BLACKWHITE
Sub1 dd FloydSteinbergA
Sub2 dd FloydSteinbergB
Sub3 dd Stucki
Sub4 dd Burkes
Sub5 dd Sierra
Sub6 dd Jarvis
Sub7 dd DOPATTERN
Sub8 dd RANDOM
Sub9 dd GETOUT





