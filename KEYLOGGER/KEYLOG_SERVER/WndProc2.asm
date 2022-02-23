;==============================================================================
; E:\Work\SubclassingThunk\2. Asm\WndProc2.asm
;
;   Subclassing Thunk (SuperClass V2) Project
;   Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
;   Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
;
;   A second attempt on WndProc thunking stub. Assembled with MASM32,
;   actually Microsoft (R) Macro Assembler Version 6.14.8444
;
; Modifications:
;
; 2002-09-28    WQW     Initial implementation based on the original
;                       WndProc.asm
;
;==============================================================================

            option casemap :none                        ;# Case sensitive
            .486                                        ;# Create 32 bit code
            .model flat, stdcall                        ;# 32 bit memory model
            .code

            WM_NCDESTROY        = 82h
            GWL_WNDPROC         = -4

start:


_wnd_proc   proc    hWin        :DWORD,
                    uMsg        :DWORD,
                    wParam      :DWORD,
                    lParam      :DWORD

            local   lReturn     :DWORD,
                    bHandled    :DWORD,
                    lInstr2     :DWORD,
                    lInstr1     :DWORD
                    

            pusha                                       ; have troubles with registers under win98
            call    _entry_point
_entry_point:
            pop     ebx                                 ; get current block ptr
            sub     ebx, offset _entry_point
            xor     eax, eax                            ; zero bHandled & lReturn
            mov     bHandled, eax
            mov     lReturn, eax
            mov     ecx, [ebx][_before_buffer_size]     ; get number of 'before' msgs
            test    ecx, ecx                            ; if anything to do
            jz      _call_orig
            cmp     ecx, -1                             ; -1 interpreted as 'all msgs'
            je      _call_before
            mov     edi, [ebx][_msg_buffer]             ; edi -> start of buffer
            mov     eax, uMsg                           ; eax -> msg number
            repne   scasd
            jne     _call_orig                          ; if not found -> call orig wndproc
_call_before:
            cmp     [ebx][_addr_ebmode], 0              ; check if in break-mode
            jz      _no_debug_check_1
            call    dword ptr [ebx][_addr_ebmode]
            cmp     eax, 2                              ; prevent re-entering VB code in break-mode
            jne     _no_debug_check_1
            mov     bHandled, 1                         ; signal debug mode -> don't even try 'after'
            jmp     _call_orig
_check_if_stopped:
            test    eax, eax                            ; if IDE 'stopped'
            jne     _no_debug_check_1
            push    [ebx][_orig_wndproc]                ; Unsubclass
            push    GWL_WNDPROC
            push    [ebx][_hwnd]
            call    dword ptr [ebx][_addr_setwindowlong]
            mov     [ebx][_sink_interface], 0           ; invalidate reference
_no_debug_check_1:
            mov     edx, [ebx][_sink_interface]         ; edx -> sink interface ptr
            test    edx, edx
            jz      _call_orig
            mov     eax, [edx]                          ; check for invalid references
            test    eax, eax
            jz      _call_orig
            and     eax, 80000000h
            jnz     _call_orig
            push    ebx                                 ; save base ptr
            lea     eax, lParam                         ; pass arguments ByRef
            push    eax
            lea     eax, wParam
            push    eax
            lea     eax, uMsg
            push    eax
            lea     eax, hWin
            push    eax
            lea     eax, lReturn
            push    eax
            lea     eax, bHandled
            push    eax
            push    edx                                 ; push 'this' ptr
            mov     eax, [edx]                          ; eax -> ptr to VTBL
            call    dword ptr [eax][20h]                ; call ISubclassingSink_Before
            pop     ebx                                 ; restore base ptr
            cmp     bHandled, 0
            jne     _return_result                      ; if handled -> return result
_call_orig:
            push    ebx                                 ; save base ptr
            push    lParam                              ; call original wndproc
            push    wParam
            push    uMsg
            push    hWin
            push    [ebx][_orig_wndproc]
            call    dword ptr [ebx][_addr_callwindowproc]
            pop     ebx                                 ; restore base ptr
            mov     lReturn, eax                        ; store result
            cmp     bHandled, 0                         ; if debug mode signalled -> return result
            jne     _return_result
            mov     ecx, [ebx][_after_buffer_size]      ; search after buffer
            test    ecx, ecx                            ; if anything to do
            jz      _return_result
            cmp     ecx, -1                             ; -1 -> special 'all msgs' case
            je      _call_after
            mov     edi, [ebx][_msg_buffer]
            mov     eax, [ebx][_before_buffer_size]     ; skip 'before' buffer
            lea     edi, [edi][4*eax]                   ; using lea instead of mul
            mov     eax, uMsg
            repne   scasd
            jne     _return_result
_call_after:
            cmp     [ebx][_addr_ebmode], 0              ; check if in break-mode
            jz      _no_debug_check_2
            call    dword ptr [ebx][_addr_ebmode]
            cmp     eax, 2                              ; prevent re-entering VB code in break-mode
            je      _return_result
_no_debug_check_2:
            mov     edx, [ebx][_sink_interface]         ; edx -> sink interface ptr
            test    edx, edx
            jz      _return_result
            mov     eax, [edx]                          ; check for invalid references
            test    eax, eax
            jz      _return_result
            and     eax, 80000000h
            jnz     _return_result
            push    ebx                                 ; save base ptr (for future enh)
            push    lParam                              ; pass arguments ByVal
            push    wParam
            push    uMsg
            push    hWin
            lea     eax, lReturn                        ; pass lReturn ByRef
            push    eax
            push    edx                                 ; push 'this' ptr
            mov     eax, [edx]                          ; eax -> ptr to VTBL
            call    dword ptr [eax][1Ch]                ; call ISubclassingSink_After
            pop     ebx                                 ; restore base ptr (for future enh)
_return_result:
            cmp     uMsg, WM_NCDESTROY                  ; if destroying -> free heap chunk
            jne     _not_destroy
            mov     [ebx][_hwnd], 0
            lea     eax, [ebx][start]                   ; free current heap chunk
            push    eax
            push    0
            push    [ebx][_process_heap]
            lea     eax, lInstr1                        ; return address
            push    eax
            mov     eax, [ebx][_not_destroy]            ; copy instruction on stack (will not be able 
            mov     lInstr1, eax                        ; to return to our just freed heap chunk!)
            mov     eax, [ebx][_not_destroy+4]
            mov     lInstr2, eax
            jmp     dword ptr [ebx][_addr_heapfree]
_not_destroy:
            popa                                        ; 61h
            mov     eax, lReturn                        ; 8Bh, 45h, 0FCh
            ret                                         ; 0C9h, 0C2h, 10h, 0

_wnd_proc   endp

            org     0190h                               ; put data block at a fixed origin

            _hwnd                   dd      ?
            _orig_wndproc           dd      ?
            _sink_interface         dd      ?
            _msg_buffer             dd      ?
            _before_buffer_size     dd      ?
            _after_buffer_size      dd      ?
            _addr_callwindowproc    dd      ?
            _addr_setwindowlong     dd      ?
            _addr_ebmode            dd      ?
            _addr_heapfree          dd      ?
            _process_heap           dd      ?


;==============================================================================
comment %
;==============================================================================

    ;#########################################################################################################################
    ;#
    ;# This code represents the model used for cSuperClass message filtered subclassing.
    ;# We assemble this code merely to discover the opcodes to use in cSuperClass.cls
    ;#
    ;# Paul_Caton@hotmail.com
    ;# 13th June 2002
    ;#
    ;# P.S. I haven't assembled since the Atari ST... That's probably self-evident.
    ;#

    .486                                ;# Create 32 bit code
    .model flat, stdcall                ;# 32 bit memory model
    option casemap :none                ;# Case sensitive
    include WndProc.inc                 ;# Macros 'n stuff

    .code

    start:

    WndProc proc    hWin    :DWORD,
                    uMsg    :DWORD,
                    wParam  :DWORD,
                    lParam  :DWORD

        LOCAL   lReturn     :DWORD
        LOCAL   lHandled    :DWORD

        jmp     TestMsgNo               ;# Jump over the *constant* code to the run-time generated message number testing code

    BeforePrevWndProc:                  ;# Jump here if we're handling this message before the previous WndProc
        mov     lReturn,0
        lea     eax,lReturn
        push    eax
        mov     lHandled,0
        lea     eax,lHandled
        push    eax
        mov     eax,88888888h           ;# Patched with ObjPtr(Owner) at run-time
        mov     ecx,eax
        mov     ecx,dword ptr [ecx]
        push    eax
        call    dword ptr [ecx+20h]     ;# Call Sub iSuperClass_Before

        cmp     lHandled,0              ;# Check to see if the user doesn't want the previous WndProc to receive this message
        jnz     Bail_1

        push    lParam                  ;# Call previous WndProc handler
        push    wParam
        push    uMsg
        push    hWin
        call    PrevWndProc             ;# PrevWndProc will be patched with the EIP relative offset to the real PrevWndProc at run-time
        ret

    AfterPrevWndProc:                   ;# Jump here if we're handling this message number after the previous WndProc
        call    PrevWndProc             ;# PrevWndProc will be patched with the EIP relative offset to the real PrevWndProc at run-time
        mov     lReturn,eax

        push    lParam                  ;# Call our handler
        push    wParam
        push    uMsg
        push    hWin
        lea     eax,lReturn
        push    eax
        mov     eax,88888888h           ;# Patched with ObjPtr(Owner) at run-time
        mov     ecx,eax
        mov     ecx,dword ptr [ecx]
        push    eax
        call    dword ptr [ecx+1Ch]     ;# Call Sub iSuperClass_After
    Bail_1:
        mov     eax,lReturn
        ret

    TestMsgNo:
        mov     eax,uMsg

        push    lParam                  ;# No matter which way we branch these parameters are going to be stacked...
        push    wParam
        push    eax
        push    hWin

        cmp     eax,0BEF00000h          ;# The compare test and jump are dynamically added at run-time
        je      BeforePrevWndProc

        cmp     eax,0AF000000h          ;# The compare test and jump are dynamically added at run-time
        je      AfterPrevWndProc

        ;#
        ;# And so on for each message number added by the user
        ;#
        ;# Unspecified messages drop thru to here...
        call    PrevWndProc             ;# We're not interested in this message number, pass to the pre-existing window proc
        ret

    PrevWndProc:

    WndProc endp

;==============================================================================
end of comment %
;==============================================================================


end start
