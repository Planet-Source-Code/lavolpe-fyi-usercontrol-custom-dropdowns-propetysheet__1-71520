;*****************************************************************************************
;** IPropertyPage hooking thunk. Assemble with nasm.
;** 
;** Callback function between VB and UserControls, 
;** subclassing the IPerPropertyPage Interface (IPPG)
;** Ref: http://msdn.microsoft.com/en-us/library/ms678432(VS.85).aspx
;** 
;** Notes:
;** 1) The cCustomPropertyDisplay class sets this thunk's edi value (See _getClient below)
;** 2) Only the IPropertyPage:Activate event is trapped
;** 3) Once above event is trapped, it subclasses the property page to this thunk
;** 	 and the thunk automatically terminates its own subclassing
;*****************************************************************************************
use32						; use 32bit registers

E_UNEXPECTED 		equ  	dWord -2147418113
GWL_WNDPROC 		equ   dWord	-4
WM_NCDESTROY		equ	dWord 130
;******************************
;Stack frame access definitions
; for the IPropertyPage:Activate thunk
%define Param3		[ebp + 48]	;::Activate thunk: property page modality value 
%define Param2          [ebp + 44]  ;::Activate: Display RECT ptr
%define Param1          [ebp + 40]  ;::Activate: hWndParent
%define iPtr            [ebp + 36]  ;::Activate thunk: calling interface pointer
; for the subclassing thunk
%define lParam		[ebp + 48]	;::Subclass thunk: SendMessage lParam							
%define wParam		[ebp + 44]	;::Subclass thunk: SendMessage wParam
%define uMsg		[ebp + 40]	;::Subclass thunk: SendMessage uMsg
%define hWnd		[ebp + 36]	;::Subclass thunk: SendMessage hWnd
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad

;***********
Align 4
;Thunk start for Subclassed PropertySite :: prevent it from showing
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	uMsg, dWord WM_NCDESTROY 
    je	_unSubclass			;handle WM_DESTROY separately; should not get here (sanity check only)
    call	_fwdMsg			;forward the message
    jmp	_Return16			;clean stack & return
_unSubclass:
    lea	eax, lReturn		;call the class that subclassed the page & let it know page is history
    push	dword eax			;pass all zero parameters so it knows it is this thunk calling back
    push	dword esi
    push	dword esi
    push	dword esi
    push	dword esi
    push	dword [edi]			;push caller's object pointer
    call	dword [edi + 0x4]		;call the function

    call	_fwdMsg			;now forward the destruction message & then unsubclass
    push	dWord [edi - 0xC]		;push prev window proc (12 bytes before edi)
    push	dWord GWL_WNDPROC		;push attribute
    push	dWord hWnd			;push hWnd
    sub	edi, 0x10 			;get SetWindowLong ptr (16 bytes before edi)
    call	near [edi]			;call the function
    jmp	_Return16
_fwdMsg:    
    push	dWord lParam		;push wnd proc's lParam
    push	dWord wParam		;push wnd proc's wParam
    push	dWord uMsg			;push wnd proc's uMsg
    push	dWord hWnd			;push wnd proc's hWnd
    push	dWord [edi - 0xC]    	;push prev wnd proc (12 bytes before edi)
    mov	eax, edi
    sub	eax, dWord 0x08		;get CallWndProc pointer (8 bytes before edi)
    call	near [eax]			;call the function
    mov	lReturn, dWord eax	;save return value
    ret

Align 4
;Thunk start for Activate		
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    call	_getClient			;call sub to get client pointer
    cmp	edi, esi			;if zero, abort
    je	_Abort16
_doActivate:				;See http reference at top of page
    lea	eax, lReturn
    push	dword eax			;push return value
    push	dword Param3		;push pointer to property page site modal value
    push	dword Param2		;push pointer to property page hosting Rect
    push	dword Param1		;push the hWndParent 
    push	dword iPtr			;push interface pointer
    push	dword [edi]			;push client object pointer
    call	dword [edi + 0x4]		;call the function
    jmp	_Return16
_Abort16:
    mov	lReturn, dword E_UNEXPECTED	;abort
_Return16:
    popad					;restore registers & places lReturn in eax
    ret	0x10				;four parameters to clean from stack

Align 4
_getClient:
nop
    xor     esi, esi		    	;Zero esi
    mov	edi, dword 012345678h	;Get Client array from VB-supplied address
    xor	eax, eax			;Zero eax
    cmp	edi, esi			;Is client array null pointer
    je	_genReturn			;Yes? > then exit & abort
    mov	eax, dWord [edi - 0xC]	;Do clients exist? (12 bytes before edi)
    cmp	eax, esi			;If not, exit & abort
    jne	_getActiveClient		;Get client to call
    xor	edi, edi			;Zero edi which is flag to abort
    ret
Align 4
_getActiveClient:
    mov	eax, dWord [edi - 0x8]	;Get object pointer of active client (if any)
    cmp	eax, esi			;if zero, then call first available client
    je	_genReturn			
    sub	edi, 0x8			;else use the active client (8 bytes before edi)
_genReturn:
    ret
