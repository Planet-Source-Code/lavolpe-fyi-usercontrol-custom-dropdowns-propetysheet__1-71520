;*****************************************************************************************
;** IPerPropertyBrowsing.asm - hooking thunk. Assemble with nasm.
;**
;** Callback function between VB and UserControls, 
;** subclassing the IPerPropertyBrowsing Interface (IPPBI)
;Ref: http://msdn.microsoft.com/en-us/library/ms678432(VS.85).aspx
;
;Notes:
; 1) The cCustomPropertyDisplay class sets this thunk's edi value (See _getClient below)
; 2) There are up to 5 class functions that can be called by this thunk
;	Only the pointer to the first is stored at edi+4. 
;	Each additional one is +32 bytes from the memory address at edi+4 (uncompiled)
;	Each additional one is +13 bytes from the memory address at edi+4 (compiled)
;*****************************************************************************************
use32						; use 32bit registers

E_NOTIMPL	equ  dWord -2147467263	; constants

;******************************
;Stack frame access definitions
%define lParam		[ebp + 48]	;if not GetDisplayString function, then: pCaCookiesOut or varOut
%define wParam          [ebp + 44]  ;depending on IPPBI request: lpDisplayName, pCaStringsOut or dwCookie
%define dispID          [ebp + 40]  ;DispatchID 
%define iPtr            [ebp + 36]  ;Ptr to IPPBI being queried
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad

;***********
Align 4
;Thunk start for GetDisplayString
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	edi, esi			;if pointer is zero, abort
    je	_Abort12
_doGetDisplayString:
    lea	eax, lReturn
    push	dword eax			;push return value
    push	dword wParam		;push return display string pointer
    push	dword dispID		;push the dispatch id
    push	dword iPtr			;push interface pointer
    push	dword [edi]			;push client object pointer
    mov	eax, dword [edi + 0x4]	;+2 functions from edi
    shl	edx, 0x1			;multiply offset*2
    add	eax, edx			;add to supplied anchor address
    call	dword eax 			;call the function
    jmp	_Return12
_Abort12:
    mov	lReturn, dword E_NOTIMPL 	;abort
_Return12:
    popad					;restore registers & places lReturn in eax
    ret	0xC				;only 3 parameters to clean from stack

Align 4
;Thunk start for GetPredefinedStrings
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	edi, esi			;if zero, abort
    je	_Abort16
_doGetPredefinedStrings:
    lea	eax, lReturn
    push	dword eax			;push return value
    push	dword lParam		;push pointer to array of longs
    push	dword wParam		;push pointer to array of strings
    push	dword dispID		;push the dispatch id
    push	dword iPtr			;push interface pointer
    push	dword [edi]			;push client object pointer
    mov	eax, dword [edi + 0x4]	;+3 function from edi
    add	eax, edx			;add one offset		(+1)
    shl	edx, 0x1			;multiply offset*2	(+2)
    add	eax, edx			;add to eax			(+3 total)
    call	dword eax			;call the function
    jmp	_Return16

Align 4
;Thunk start for GetPredefinedValue
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	edi, esi			;if zero, abort
    je	_Abort16
_doGetPredefinedValue:
    lea	eax, lReturn
    push	dword eax			;push return value
    push	dword lParam		;push pointer to array of variant return value
    push	dword wParam		;push pointer to cookie
    push	dword dispID		;push the dispatch id
    push	dword iPtr			;push interface pointer
    push	dword [edi]			;push client object pointer
    mov	eax, dword [edi + 0x4]	;+4 functions from edi
    shl	edx, 0x2			;multiply offset*4
    add	eax, edx			;add to eax
    call	dword eax			;call the function
    jmp	_Return16
_Abort16:
    mov	lReturn, dword E_NOTIMPL	;abort
_Return16:
    popad					;restore registers & places lReturn in eax
    ret	0x10				;four parameters to clean from stack

Align 4
;Thunk start for MapToPropertyPage
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	edi, esi			;if pointer is zero, abort
    je	_Abort12
_doMapProperty:
    lea	eax, lReturn
    push	dword eax			;push return value
    push	dword wParam		;push return CLSID pointer
    push	dword dispID		;push the dispatch id
    push	dword iPtr			;push interface pointer
    push	dword [edi]			;push client object pointer
    mov	eax, [edi + 0x4]		;+5 functions from edi
    add	eax, edx			;add one offset		(+1)
    shl	edx, 0x2			;multiply offset*4	(+4)
    add	eax, edx			;add to eax			(+5 total)
    call	dword eax			;call the function
    jmp	_Return12

Align 4
_getClient:
nop						;alignment to get 012345678h on dWord boundary
    xor     esi, esi		    	;Zero esi
    mov	edi, dword 012345678h	;Get Client array from VB-supplied address
    xor	eax, eax			;Zero eax
    cmp	edi, esi			;Is client array null pointer
    je	_genReturn			;Yes? > then exit & abort
    mov	eax, dWord [edi - 0x4]	;Does clients exist?
    cmp	eax, esi			;If not, exit & abort
    jne	_getClientFromHost	;See if more than one exists
    xor	edi, edi			;Zero edi which is flag to abort
    ret
Align 4
_getClientFromHost:
nop						;alignment to get 012345678h on dWord boundary
nop
nop
    mov	edx, dWord 012345678h	;function offset to other functions: either 32 or 13
    cmp	eax, dword 0x1		;Just one client?
    je	_genReturn			;Return then; else ask client for the correct client
    push	edx
    lea	eax, lReturn		
    push	dword eax			;push our return value
    push	dword edi			;push memory location for clients
    push	dword iPtr			;push the interface pointer
    push	dword [edi]			;get reference to Class
    mov	eax, dWord [edi + 0x4]
    add	eax, edx			;get next function pointer
    call	dWord eax			;call the helper function
    mov	edi, dword lReturn	;transfer passed pointer to edi
    pop	edx
_genReturn:
    ret
