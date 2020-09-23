VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomPropertyDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Please take about 2 minutes and read the comments below. It will help you understand
' what this class does and cannot do, the basics of how it works. The interface
' manipulation is pretty advanced stuff, so I have added a ton of remarks in almost
' every function/sub within this class, and also the IPropertyBrowserEvents class.



' This class uses some VTable hacks to achieve its purpose. Also, since this class' only
' purpose is to "pretty up" the property sheet during design time, it should not be
' invoked during runtime.  The Attach function has a parameter that you pass to tell the
' class whether or not the usercontrol's parent is in design time or runtime.

' ========== WARNINGS  WARNINGS  WARNINGS  WARNINGS  WARNINGS  WARNINGS ============
' You should NOT walk through this class or its events within your usercontrol. The class
' basically subclasses an interface class called IPerPropertyBrowsing. This requires
' subclassing that interface during design time and, with any subclassing, it is not
' safe to step thru the code while subclassing is in play. You can step thru and debug
' the rest of your project without any problems.

' If this class is running in a usercontrol that is not compiled and you make changes to
' that usercontrol while it is also displayed on a form, you may crash. This is because
' the Terminate events are not fired when VB destroys this class and/or usercontrol. The
' Terminate events allow unsubclassing to occur. I have coded around this problem the best
' I can, but potential of crash still exists. Therefore, there is a protection property
' that you must set to true to allow this to run on uncompiled usercontrols.
' That property is IgnoreIDESafety.

' You have been warned. This class should be completely safe in compiled OCXs.
' ===============================================================================

' ----------- KNOWN ISSUES  KNOWN ISSUES  KNOWN ISSUES  KNOWN ISSUES ------------
' 1. With VB, when multiple controls are selected, the property sheet will show
'    most, if not all, of the usercontrol's properties. This allows the user to
'    change a property and it affects all controls selected.  This behavior
'    is not performed when a property is implemented via the IPerPropertyBrowser.
'    All properties that are not implemented are shown as expected
'
' 2. With VB, when a dropdown list is provided containing a list to select from
'    and user double clicks the property name, VB will automatically select the
'    next item in the list in a circular fashion. Using this class can fail to
'    allow that behavior unless the following is enforced by you. Whenever, you
'    implemenet a property within this class... You provide the class with the
'    dropdown listing, and when prompted to set the display text for the property,
'    you MUST use the exact same text in the dropdown listing you provided for
'    the item that was selected.  Only if it matches exactly does this behavior
'    continue as expected.
'
' 3. When wanting to display a dialog box, custom modal form, etc, for a property,
'    some minor issues are noted: a) You may get a quick "flash" of a subclassed
'    property page closing after your dialog box is closed b) Cannot override
'    properties that are set as stdPicture, stdFont or OLE_COLOR. VB will display
'    what it wants to for these, no matter what we tell it. If you absolutely want
'    your own dialog, declare such properties as Variant or Long.
'
' 4. There is a noticable delay when clicking on an uncompiled usercontrol and
'    when the property sheet is updated with that control's properties. There is
'    no delay when the usercontrol is compiled to an ocx.
' ------------------------------------------------------------------------------

' This class can add a bit more professionalism to finished usercontrols.
' With the class, you can completely control what text is displayed in the property sheet
' for properties that get their values from dropdown lists (i.e., enumerations)
' Though one could supply unicode text, the property sheet in VB does not support unicode.
' You can also make your properties expose ellipse (...) buttons so you can display
' your own dialog or modal form intead of a property sheet.

' Here are some generalized examples. More exist in IPropertyBrowserEvents class.

' The text displayed in a dropdown box containing an enumeration
'   Example: Your Enum is 0 - vbLeftJustify, 1 - vbRightJustify, 2 - vbCenter
'       You could display this instead: 0 - Left Aligned, 1 - Right Aligned, 2 - Centered
'       Or possibly this:  0 - Left, 1 - Right, 2 - Center
' On the fly, dynamic, "enumerations".
' For example, you can provide a list of available drives or a list of controls on the parent
' Without similar implementations, this was only possible with the use of property pages
' See the IPropertyBrowserEvents class for coded examples

' So how do you use this class?
' 1. Declare in your usercontrol's (UC) declarations section:
'       Dim WithEvents myCustomProps As cCustomPropertyDisplay
' 2. Add this to your declarations section
'       Implements IPropertyBrowserEvents
' 3. Add this somewhere in your UC
'        Private Sub ImplementProperties()
'            Set myCustomProps = New cCustomPropertyDisplay
'            If myCustomProps.Attach(Me, Ambient.UserMode) Then
'                ' for each custom property, add it to the class collection
'                myCustomProps.AddProperty Me, "Drives"
'                myCustomProps.AddProperty Me, "Alignment"
'            Else
'                Set myCustomProps = Nothing
'            End If
'        End Sub
' 4. In your UC's InitProperties event, add this line of code
'       Call ImplementProperties
' 5. In your UC's ReadProperties event, add this line of code
'       Call ImplementProperties
' 6. In you UC's Terminate event, add this line of code
'       If Not myCustomProps Is Nothing Then myCustomProps.Detach
' 7. Respond to the following new events contained in your UC
'       IPropertyBrowserEvents_FormatPropertyDisplay
'       IPropertyBrowserEvents_FormatPropertyEnum
'       IPropertyBrowserEvents_SetEnumPropertyValue
'       IPropertyBrowserEvents_ShowCustomPropertyPage
' see IPropertyBrowserEvents class for more info & examples

' There is one addiitonal function provided that is not associated directly with the usercontrol
' and is provided for convenience only. PropertyPageClose (see that function for comments)
' That function allows you to close a property page, via code, from within the property page


' PROGRAMMING NOTES
' A vast portion of this class is simply to track & maintain arrays used for the thunk. Why use a thunk & why
' manage memory arrays in memory vs in a module or in the class itself?  The reason is kinda simple.
' VB will use the same interface for multiple controls of the same type.  If your usercontrol hacks the vTable
' so it can receive events, what happens when the next usercontrol is loaded?  It hacks the same vTable and
' all of the previous usercontrol's events are fired to last usercontrol that hacked the vTable.  This means, that,
' in effect, all the properties of previous usercontrols can change when the last one's properties are changed.
' To make matters worse, compiled OCXs have their own vTable, though many copies of the control can use the
' same VTable.  But in IDE design, any uncompiled usercontrols, whether the same usercontrol or different,
' can use the same VTable.

' Therefore, a memory thunk is used that exists in project memory and that can call any usercontrol in the project
' and all usercontrols know about it and have access to it. And since we can't allow the usercontrol references
' and pointers to exist in a single class (should that class be destroyed), we have to have them float in memory.
' Managing memory arrays was fun, but is a bit code intensive. Note: regarding the thunks. The thunks call
' functions within this class via Paul Caton's great zAddressOf routines. To prevent having to store the address
' of up to 6 functions per class instance, the thunks use offsets from a known address. When uncompiled, this
' offset, within class objects, appears to be 32 bytes between functions. When compiled, it drops to 13 bytes.
' Keep this in mind should you want to tweak the asm thunks.

' http://www.vbaccelerator.com/home/vb/Code/Controls/UserControl_OLE_Extensions/IPerPropertyBrowse/article.asp
' The example on vbAccelerator uses a DLL to prevent these problems that I overcame via thunks & managing memory
' arrays. The vbAccelerator post is a good bit of coding. I wanted to try and do this without using an Active-X
' DLL and TLBs. I have managed to replicate vbAccelerator's code exactly and also improved/fixed some issues
' their DLL experiences.

' This class makes heavy use of the DispCallFunc API which enables coders low-level interface manipulation. It isn't
' exactly easy to use and there is very little documentation on it, but if you have an interface pointer, you have
' full control over it with minimal effort. Researching the Intefaces on MSDN is a must though, crashing is easy :)

' Is this a better solution than vbAccelerator's version?  Nope, just a different one that uses no additional dependencies
' If you do decide to use this code, you must add the 2 classes to your usercontrol project.
' Obviously the form & usercontrol provided in this project are soley to enable experimenting.


' APIs primarily for setting up memory arrays and communicating via interface pointers
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal ppv As Long, ByVal oVft As Long, ByVal cc As Long, ByVal rtTYP As VbVarType, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef fuReturn As Variant) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As Long
Private Declare Function CoTaskMemRealloc Lib "ole32.dll" (ByVal pv As Any, ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SysAllocString Lib "oleaut32.dll" (ByVal pOlechar As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SAFEARRAYBOUND
End Type
Private Type CAWORDOLESTR
    Count As Long
    DataPtr As Long
End Type
Private Const E_POINTER As Long = &H80004003
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_OUTOFMEMORY As Long = &H8007000E
Private Const E_NOINTERFACE As Long = &H80004002

' APIs used to create & maintain mapped file
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByRef lpBaseAddress As Any) As Long
Private Const PAGE_READWRITE As Long = &H4
Private Const PAGE_EXECUTE_READWRITE& = &H40&
Private Const ERROR_ALREADY_EXISTS As Long = 183&
Private Const SECTION_MAP_WRITE As Long = &H2
Private Const SECTION_MAP_READ As Long = &H4
Private Const CC_STDCALL As Long = 4&

' APIs used to control display of a property page
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Long) As Long
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_SETREDRAW As Long = &HB

'-Callback declarations for Paul Caton thunking magic----------------------------------------------
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
'-------------------------------------------------------------------------------------------------

' -------- Following used in PropertyPageClose function
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type
Private Const VK_RETURN As Long = &HD
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_ESCAPE As Long = &H1B

' -------- Local constants, structures and class variables
Private Const IUnknownQueryInterface As Long = 0&   ' IUnknown vTable offset to Query implemented interfaces
Private Const IUnknownAddRef As Long = 4&           ' IUnknown vTable offset to increment reference count
Private Const IUnknownRelease As Long = 8&          ' IUnkownn vTable offset to decrement reference count

Private Const ThunkVTableOffset As Long = &HCC      ' offset in thunk to store client array pointer (See CreateThunk)
Private Const ThunkVTableOffsetPPG As Long = &H9C   ' offset in thunk to store client array pointer (See CreateThunk)
Private Const ThunkGDSOffset As Long = &H0&         ' offset in thunk where GetDisplayString is handled
Private Const ThunkGPDSOffset As Long = &H34        ' offset in thunk where GetPredefinedStrings is handled
Private Const ThunkGPDVOffset As Long = &H64        ' offset in thunk where GetPredefinedValue is handled
Private Const ThunkMPTPOffset As Long = &H9C        ' offset in thunk where MapPropertyToPage is handled
Private Const ThunkActivateOffset As Long = &H64    ' offset in thunk where IPropertyPage:Activate is handled

Private Enum PropType                               ' property type with regards to this class
    ptype_Enum = 0&
    ptype_Dialog = 1&
    ptype_LockedDialog = 3&
End Enum
Private Type UserDispActionStruct                   ' used to store implemented property info
    DispID As Long                                  ' IDispatch ID of the property
    Name As String                                  ' Actual property name
    pType As PropType                               ' 0=enumeration-type, 1=custom dialog, 2=custom dialog w/locking display
End Type
Private Type DispIDStruct                           ' Collection of property info
    Count As Long
    Params() As UserDispActionStruct
End Type

' IPPB=IPerPropertyBrowsing, IPPG=IPropertyPage
Private Type VTableArrayStruct      ' (one structure per unique VTable that is subclassed) :: 28 bytes, and may have an additonal 4 byte counter.
    VTable As Long                  ' the subclassed interface's VTable address
    ClientPtr As Long               ' pointer to the VTableArrayClientStruct array
    ThunkPtr As Long                ' address where thunk resides
    origFunction1 As Long           ' original VTable function address: for IPPB; IPerPropertyBrowsing_GetDisplayString, for IPPG; IPropertyPage_Activate
    origFunction2 As Long           ' original VTable function address: for IPPB; IPerPropertyBrowsing_GetPredefinedStrings, for IPPG; IPPG reference
    origFunction3 As Long           ' original VTable function address: for IPPB; IPerPropertyBrowsing_GetPredefinedValue, for IPPG; not used
    origFunction4 As Long           ' original VTable function address: for IPPB; IPerPropertyBrowsing_MapToPropertyPage, for IPPG; not used
End Type
Private Type VTableArrayClientStruct ' (one structure per class instance) :: 12 bytes and may have an additional 4-24 bytes as shown below
    MePointer As Long               ' this class' object pointer
    LocalVTableOffset As Long       ' this class' IPropertyPage:Activate pointer; additional functions are offset by +/- 32-13 bytes each
    InterfacePtr As Long            ' IPPB: reference to the IPerPropertyBrowsing interface, IPPG: not used
End Type
' Client array structure.
' IPPB : 4 bytes       12 bytes     12 bytes
'        ArrayCount    1st client   2nd client  etc
'                      ^^ ptr stored in thunk
' IPPG : 4 bytes       4 bytes      4 bytes          4 bytes      4 Bytes          4 bytes              12 bytes    12 bytes
'        SetWinLongPtr PrevWndProc  CallWndProcPtr   ArrayCount   ActiveClientObj  ActiveClientFuncPtr  1st client  2nd client  etc
'        ^^ set by IPropertyPage:Activate ^^                       ^^ Set by MapPropertyToPage ^^       ^^ ptr stored in thunk

' cached pointer references
Private c_UCIPEPointer As Long          ' pointer to calling usercontrol's IPropertyBrowserEvents object
Private c_InterfaceIPPB As Long         ' pointer to calling usercontrol's IPerPropertyBrowser interface
Private c_localVTableOffset As Long     ' function pointer to this class' IPropertyPage:Activate
Private c_FileHandle As Long            ' pointer to shared mapped memory file (all implemented usercontrols share this file)
' the mapped file structure follows :: 12 bytes total, each are pointers to a VTableArrayStruct
' 4 bytes                            4 bytes                              4 bytes
' vTables() array for compiled ocx   vTables() array for uncompiled ocx   vTable for hacked property page

' configuration variables
Private c_UncompiledOCX As Long         ' flag indicating compiled class or not & also c_FileHandle offset: 0=compiled, 4=uncompiled
Private c_RunUnsafeInIDE As Boolean     ' IDE safety override flag: See IgnoreIDESafety property
' collection
Private c_DispIDCol As DispIDStruct     ' collection of implemented dispatch IDs and related information
Private c_MappedDispID As Long          ' disposition ID of last mapped property (See IPerPropertyBrowsing_MapToPropertyPage)

Public Property Let IgnoreIDESafety(bIgnore As Boolean)
    ' as mentioned in the lead remarks at very top, this class
    ' has potential of crashing your project when running it with
    ' an uncompiled usercontrol. By default it is not allowed
    ' to run unless 1) the usercontrol is compiled, or 2) this
    ' property is set to true before Attach is called
    c_RunUnsafeInIDE = bIgnore
End Property

Public Function Attach(theUsercontrol As Object, ByVal AmbientUserMode As Boolean) As Boolean

    ' Must be called before AddProperty function is called.
    ' This function associates a IPerPropertyBrowsing Interface with the usercontrol and then
    ' tweaks the interface's VTable to send 3 of the interfaces's functions to this class vs VB
    ' The Detach property must be called in the usercontrol's terminate event
    
    ' The AmbientUserMode must be passed as Ambient.UserMode
    ' When UserMode=True, then usercontrol is in runtime & modifying the VTable is not performed
    ' because this class is for container's design time only, not runtime
    
    ' Function fails on these conditions
    ' 1) Attach already called
    ' 2) AmbientUserMode=False and IgnoreIDESafety=False
    ' 3) AmbientUserMode=True (should fail, you do not want to subclass at runtime)
    ' 4) Failed to obtain an interface pointer for either IPropertyPage or IPerPropertyBrowsing (unlikely)
    ' 5) Failed to create needed array memory (usually less than 1kb per project). Low memory issues
    
    If AmbientUserMode = False Then ' else no need to subclass the Interfaces if running in compiled mode
        
        If c_FileHandle Then Exit Function ' already attached
        
        Dim iUCevents As IPropertyBrowserEvents
        Debug.Assert IsUnCompiled()                     ' running in IDE?
        If c_RunUnsafeInIDE = False Then                ' IgnoreIDESafety set? If not...
            If c_UncompiledOCX Then Exit Function       ' yes, abort because IgnoreIDESafety property not set
        End If
        
        On Error Resume Next
        Set iUCevents = theUsercontrol                  ' test if calling object Implements IPropertyBrowserEvents
        If Err Then
            Err.Clear                                   ' usercontrol failed to add Implements IPropertyBrowserEvents
            Exit Function
        End If
        c_UCIPEPointer = ObjPtr(iUCevents)              ' cache pointer to calling object's IPropertyBrowserEvents object
        
        Dim varRtn As Variant, pUC As Long
        Dim GUID(0 To 3) As Long
        Dim pvTypes(0 To 1) As Integer, pvPtrs(0 To 1) As Long, pValues(0 To 1) As Variant
        Const IID_IPerPropertyBrowsing As String = "{376BD3AA-3845-101B-84ED-08002B2EC713}" ' GUID for IPerPropertyBrowsing
        
        On Error GoTo CATCH_EXCEPTION
        If ConvertGUIDtoArray(IID_IPerPropertyBrowsing, GUID) Then
            pUC = ObjPtr(theUsercontrol)
            pvTypes(0) = vbLong: pValues(0) = VarPtr(GUID(0)): pvPtrs(0) = VarPtr(pValues(0))
            pvTypes(1) = vbLong: pValues(1) = VarPtr(c_InterfaceIPPB): pvPtrs(1) = VarPtr(pValues(1))
            ' See if the usercontrol implements IPerPropertyBrowsing & it should if it has any public variables or methods
            Call DispCallFunc(pUC, IUnknownQueryInterface, CC_STDCALL, vbLong, 2&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
            If c_InterfaceIPPB Then ' we have IPerPropertyBrowsing interface
                ' QueryInterfaces ups the ref count, release the reference
                Call DispCallFunc(c_InterfaceIPPB, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
                Attach = AddClient_IPerPropertyBrowsing()                    ' call workhorse for adding clients, thunks & managing memory arrays
            End If
        End If
    End If
CATCH_EXCEPTION:
    If Err Then Err.Clear
End Function

Public Function Detach() As Boolean

    ' This function removes a client from the thunks client list and optionally...
    '   -- destroys/resizes the memory array
    '   -- modifies the thunk's client pointer
    '   -- restores the VTable's function pointers
    '   -- destroys/updates the mapped memory file
    
    ' Function can fail for these conditions
    ' 1) Attach was never called

    If c_InterfaceIPPB = 0& Then Exit Function
    If c_FileHandle = 0& Then Exit Function

    Dim Clients() As VTableArrayClientStruct, vTables() As VTableArrayStruct
    Dim tSA As SafeArray, vTablePtr As Long
    Dim cSA As SafeArray, ClientPtr As Long
    Dim VTable As Long, hView As Long, varRtn As Variant
    Dim vCount As Long, cCount As Long, vIndex As Long, cIndex As Long, lRtn As Long
    
    Const MEM_RELEASE As Long = &H8000&                        'Release allocated memory flag

    hView = MapViewOfFile(c_FileHandle, SECTION_MAP_WRITE, 0&, 0&, 0&)  ' map the memory file
    CopyMemory vTablePtr, ByVal hView + c_UncompiledOCX, 4&         ' get array ptr to VTable array
    CopyMemory vCount, ByVal vTablePtr, 4&                          ' get the number of VTable array entries
    CopyMemory VTable, ByVal c_InterfaceIPPB, 4&                        ' get the vTable pointer for this uc's IPPB interface
    
    OverlayArrayOnMemory tSA, VarPtrArray(vTables), vTablePtr + 4&, 28&, vCount
    vIndex = FindVTable(vTables(), VTable)                          ' find the vTable for this interface
    If vIndex < vCount Then                                          ' else not found, should never happen
        ClientPtr = vTables(vIndex).ClientPtr                       ' cache the client pointer
        CopyMemory cCount, ByVal ClientPtr, 4&                      ' get the client count & overlay array on memory address
        OverlayArrayOnMemory cSA, VarPtrArray(Clients), ClientPtr + 4&, 12&, cCount
        cIndex = FindClient(Clients(), c_InterfaceIPPB)                 ' search for this client
        If cIndex < cCount Then                                     ' else client not found, should never happen
            cCount = cCount - 1                                     ' decrement count
            If cCount = 0& Then                                     ' removing last client
                CoTaskMemFree ClientPtr                                         ' free array memory
                SetVTableEntry VTable + 12&, vTables(vIndex).origFunction1  ' replace vTable pointers
                SetVTableEntry VTable + 20&, vTables(vIndex).origFunction2
                SetVTableEntry VTable + 24&, vTables(vIndex).origFunction3
                SetVTableEntry VTable + 16&, vTables(vIndex).origFunction4
            Else                                                    ' shift client array left, deleting the detached client
                If cIndex < cCount Then CopyMemory Clients(cIndex), Clients(cIndex + 1), (cCount - cIndex) * 12&
            End If
            OverlayArrayOnMemory cSA, VarPtrArray(Clients), 0&, 0&, 0&  ' remove overlay, no longer needed
            If cCount Then
                lRtn = CoTaskMemRealloc(ClientPtr, cCount * 12& + 4&)   ' redim preserve the clients array
                If lRtn Then                                            ' else error out of memory, we can't redim the array
                    vTables(vIndex).ClientPtr = lRtn                    ' set vTable's pointer & thunk address client ref
                    CopyMemory ByVal vTables(vIndex).ThunkPtr + ThunkVTableOffset, lRtn + 4&, 4&
                    CopyMemory ByVal lRtn, cCount, 4&                   ' update client count
                End If
            Else                                                        ' removed last client
                VirtualFree vTables(vIndex).ThunkPtr, 0&, MEM_RELEASE   ' release the thunk
                vCount = vCount - 1                                     ' decrement count
                If vCount = 0& Then                                     ' removing last vtable reference; free array memory
                    CoTaskMemFree vTablePtr
                    CopyMemory ByVal hView + c_UncompiledOCX, vCount, 4&
                Else                                                    ' shift array & then redim preserve it
                    If vIndex < vCount Then CopyMemory vTables(vIndex), vTables(vIndex + 1), (vCount - vIndex) * 28&
                    lRtn = CoTaskMemRealloc(vTablePtr, vCount * 28& + 4&)
                    If lRtn Then CopyMemory ByVal hView + c_UncompiledOCX, lRtn, 4& ' resized, update the mapped memory file
                End If
            End If
            OverlayArrayOnMemory tSA, VarPtrArray(vTables), 0&, 0&, 0&  ' remove overlay, no longer needed
            Detach = True                                               ' return success
        Else
            OverlayArrayOnMemory cSA, VarPtrArray(Clients), 0&, 0&, 0&  ' remove overlays, no longer needed
            OverlayArrayOnMemory tSA, VarPtrArray(vTables), 0&, 0&, 0&
        End If
    Else
        OverlayArrayOnMemory tSA, VarPtrArray(vTables), 0&, 0&, 0&      ' remove overlay, no longer needed
    End If
    
    ' check to see if IPropertyPage was created...
    CopyMemory vTablePtr, ByVal hView + 8&, 4&                          ' get hacked property page vTable() array pointer
    If vTablePtr Then
        ReDim vTables(0 To 0)                                       ' copy vTable array from memory
        CopyMemory vTables(0), ByVal vTablePtr, 28&
        CopyMemory cCount, ByVal vTables(0).ClientPtr + 12&, 4&     ' get client count & overlay local array on memory array
        OverlayArrayOnMemory tSA, VarPtrArray(Clients), vTables(0).ClientPtr + 12&, 12&, cCount + 1
        ' find us in the client array
        lRtn = ObjPtr(Me)
        For cIndex = 1 To cCount
            If Clients(cIndex).MePointer = lRtn Then Exit For
        Next
        If cIndex <= cCount Then                                    ' we are a client, remove and reallocate array
            cCount = cCount - 1                                     ' decrement client count
            If cCount = 0& Then                                     ' we were last client
                VirtualFree vTables(0).ThunkPtr, 0&, MEM_RELEASE    ' release thunk
                CoTaskMemFree vTables(0).ClientPtr                  ' release client array
                CoTaskMemFree vTablePtr                             ' release vTable array
                CopyMemory ByVal hView + 8&, 0&, 4&                 ' remove reference from mapped file
                SetVTableEntry vTables(0).VTable + 16&, vTables(0).origFunction1
                ' unreference the PropertyPage we created in AddClient_IPropertyPage
                Call DispCallFunc(vTables(0).origFunction2, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
            Else                                                    ' shift clients left
                If cIndex < cCount Then CopyMemory Clients(cIndex), Clients(cIndex + 1), (cCount - cIndex) * 12&
                Clients(0).MePointer = cCount                       ' update client count & resize array
                lRtn = CoTaskMemRealloc(vTables(0).ClientPtr, cCount * 12& + 24&)
                If lRtn Then                                        ' couldn't shorten array? hmmm, memory issues
                    CopyMemory ByVal vTablePtr + 4&, lRtn, 4&       ' otherwise, update pointers in thunk & vTable structure
                    CopyMemory ByVal vTables(0).ThunkPtr + ThunkVTableOffsetPPG, lRtn + 24&, 4&
                End If
            End If
        End If
        OverlayArrayOnMemory tSA, VarPtrArray(Clients), 0&, 0&, 0&
        Erase vTables()
    End If
    If hView Then UnmapViewOfFile hView                                 ' unmap file, if not already done
    CloseHandle c_FileHandle                                                ' unreference the shared memory file
    c_InterfaceIPPB = 0&                                                    ' reset local variables
    c_FileHandle = 0&
End Function

Public Function AddProperty(theUsercontrol As Object, ByVal PropertyName As String, _
                                                Optional Enumeration As Boolean = True, _
                                                Optional LockIDEdisplay As Boolean = False) As Boolean
    
    ' Must call Attach function first.
    ' Once the class is attached to a usercontrol, call this function to add all properties that
    ' you want to display custom dropdown enumeration text
    
    ' Parmeters
    ' theUserControl :: Always pass as Me
    ' PropertyName :: the property to be implemented
    ' Enumeration ::
    '   If True, the property has values displayed in a dropdown combobox
    '   If False, the property gets value from custom dialog or form you display
    ' LockIDEDisplay :: applicable only if Enumeration=False
    '   If True, a user cannot add property directly into the property sheet, must use ellipsis button (...)
    '       If true, then you should also respond to the IPropertyBrowserEvents_FormatPropertyDisplay event
    '       However, if you fail to respond to that event or set its Cancel parameter to True, the display is unlocked
    '   If False, user can type/paste values into the property sheet, prevents unnecessary calls to your usercontrol
    
    ' Function fails on these conditions
    ' 1) Attach not called yet
    ' 2) Property already added or Property doesn't exist: PropertyName is invalid
    ' 3) Failed to obtain an interface pointer for either IPropertyPage or IPerPropertyBrowsing (unlikely)
    ' 4) Failed to create needed array memory (usually less than 1kb per project). Low memory issues
    ' 5) Coding error on my part -- these can be fixed though
    If c_FileHandle = 0& Then Exit Function ' must call Attach first
    
    Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}" ' GUID for IDispatch
    Const IDispatchIDsOfNames As Long = 20& ' offset from IDispatch VTable
    
    Dim varRtn As Variant, DispID As Long
    Dim ptrIInterface As Long, pUnk As Long
    Dim GUID(0 To 3) As Long, CLSID(0 To 3) As Long
    Dim pvTypes(0 To 4) As Integer, pvPtrs(0 To 4) As Long, pValues(0 To 4) As Variant
        
    On Error GoTo CATCH_ERROR
    If ConvertGUIDtoArray(IID_IDispatch, GUID) Then
        pUnk = ObjPtr(theUsercontrol)
        ' Call UC's QueryInterface to get the IDispatch interface pointer
        pvTypes(0) = vbLong: pValues(0) = VarPtr(GUID(0)): pvPtrs(0) = VarPtr(pValues(0))
        pvTypes(1) = vbLong: pValues(1) = VarPtr(ptrIInterface): pvPtrs(1) = VarPtr(pValues(1))
        ' HRESULT QueryInterface(REFIID iid, void ** ppvObject);
        Call DispCallFunc(pUnk, IUnknownQueryInterface, CC_STDCALL, vbLong, 2&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
        If ptrIInterface Then ' we have IDispatch interface, query the ID of the passed method
            Erase GUID
            pvTypes(0) = vbLong: pValues(0) = VarPtr(GUID(0)):        pvPtrs(0) = VarPtr(pValues(0))
            pvTypes(1) = vbString: pValues(1) = VarPtr(PropertyName): pvPtrs(1) = VarPtr(pValues(1))
            pvTypes(2) = vbLong: pValues(2) = 1&:                     pvPtrs(2) = VarPtr(pValues(2))
            pvTypes(3) = vbLong: pValues(3) = 0&:                     pvPtrs(3) = VarPtr(pValues(3))
            pvTypes(4) = vbLong: pValues(4) = VarPtr(DispID):         pvPtrs(4) = VarPtr(pValues(4))
            ' HRESULT GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR*  rgszNames, unsignedint cNames, LCID lcid, dispID FAR * rgDispId);
            Call DispCallFunc(ptrIInterface, IDispatchIDsOfNames, CC_STDCALL, vbLong, 5&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
            ' QueryInterfaces ups the ref count, release the reference
            Call DispCallFunc(ptrIInterface, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
            If DispID <> -1& Then                               ' add the DispatchID to our local array
                If FindDispID(DispID, False) Then Exit Function ' abort; already added this property
                If Enumeration = False Then                     ' validate we can do this before we implement the property
                    If AddClient_IPropertyPage() = False Then Exit Function
                End If
                DispID = FindDispID(DispID, True)               ' add property to our collection
                c_DispIDCol.Params(DispID).Name = PropertyName  ' cache property name for user callbacks
                If Enumeration Then                             ' set the property type used for interface callbacks
                    c_DispIDCol.Params(DispID).pType = ptype_Enum
                Else
                    If LockIDEdisplay Then c_DispIDCol.Params(DispID).pType = ptype_LockedDialog Else c_DispIDCol.Params(DispID).pType = ptype_Dialog
                End If
                AddProperty = True
            End If
        End If
    End If
CATCH_ERROR:
If Err Then Err.Clear
End Function

Public Function PropertyPageClose(thePropertyPage As Object, ByVal thePropertyPageHwnd As Long, Optional ByVal viaApplyButton As Boolean = True) As Boolean

    ' Should you ever need to close a property page from within the property page, via code...
    ' thePropertyPage parameter is passed as:  Me
    ' thePropertyPageHwnd is passed as:  PropertyPage.hWnd
    ' viaApplyButton if True will call the property page's Apply event, regardless if the page's Changed/Dirty property is True
    ' Sample Call:  PropertyPageClose Me, PropertyPage.hwnd, False
    
    ' The function will fail if viaApplyButton=True and the property page failed to apply the changes.
    ' Note: This cannot be called in the propertypage's Initialize event, VB will override and display the page anyway
    
    Const IID_IPropertyPage As String = "{B196B28D-BAB4-101A-B69C-00AA00341D07}"
    Const IPropPageApplyOffset As Long = 44& ' 12th VTable entry
    Const IPropPageTranslateKeyOffset As Long = 52& ' 14th VTable entry
    
    Dim pMSG As Msg
    Dim pUnk As Long
    Dim GUID(0 To 3) As Long, varRtn As Variant, ptrIPropPage As Long
    Dim pvTypes(0 To 1) As Integer, pvPtrs(0 To 2) As Long, pValues(0 To 1) As Variant
    
    On Error GoTo CATCH_EXCEPTION
    If ConvertGUIDtoArray(IID_IPropertyPage, GUID) Then
        ' Call QueryInterface to see if the IPropertyPage interface is used
        ' Note: Per MSDN, IPropertyPage must be used if property page is displayed, using IPropertyPage2 is optional
        pvTypes(0) = vbLong: pValues(0) = VarPtr(GUID(0)): pvPtrs(0) = VarPtr(pValues(0))
        pvTypes(1) = vbLong: pValues(1) = VarPtr(ptrIPropPage): pvPtrs(1) = VarPtr(pValues(1))
        pUnk = ObjPtr(thePropertyPage)
        Call DispCallFunc(pUnk, IUnknownQueryInterface, CC_STDCALL, vbLong, 2&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
        If ptrIPropPage Then
            ' It is used, see if user wants the apply method called too?
            Call DispCallFunc(ptrIPropPage, IUnknownAddRef, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
            If viaApplyButton Then ' abort if Apply failed
                If DispCallFunc(ptrIPropPage, IPropPageApplyOffset, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn) <> 0& Then
                    Call DispCallFunc(ptrIPropPage, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
                    Exit Function
                End If
            End If
            ' Send the property page an Escape Key to close
            pMSG.hWnd = thePropertyPageHwnd
            pMSG.Message = WM_KEYDOWN
            pMSG.wParam = VK_ESCAPE
            pMSG.lParam = -1
            pvTypes(0) = vbLong: pValues(0) = VarPtr(pMSG): pvPtrs(0) = VarPtr(pValues(0))
            Call DispCallFunc(ptrIPropPage, IPropPageTranslateKeyOffset, CC_STDCALL, vbLong, 1&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
            ' QueryInterfaces ups the ref count, release the reference
            Call DispCallFunc(ptrIPropPage, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
            PropertyPageClose = True
        End If
    End If
CATCH_EXCEPTION:
    If Err Then Err.Clear
    If ptrIPropPage Then Call DispCallFunc(ptrIPropPage, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
End Function

Private Function AddClient_IPerPropertyBrowsing() As Boolean

    ' This function has multiple purposes and is the workhorse for Attach
    ' 1) Creates a shared mapped memory file
    ' 2) Creates a thunk for each new vTable that we may be subclassing
    ' 3) Creates/resizes memory arrays used by the thunk
    ' 4) Hacks the vTable, redirecting its functions to this class
    
    Dim hView As Long, bNew As Boolean, bError As Boolean
    Dim theTables() As VTableArrayStruct, vCount As Long
    Dim theClients() As VTableArrayClientStruct, cCount As Long
    Dim arrayPtr As Long, tSA As SafeArray, cSA As SafeArray, Index As Long
    Dim theClient As Long, VTable As Long
    Dim vTableData As VTableArrayStruct
    
    CopyMemory VTable, ByVal c_InterfaceIPPB, 4&   ' VTable address
    ' create an object we can use that will disappear when IDE closes & that can be shared by all usercontrols, compiled or not
    ' ensure to make it unique to just our Thread or Process. Mapped memory files can be shared across other processes, and if
    ' we have uc in VB while uc also in MSaccess, the vTable & Client array pointers will be invalid memory addresses in the
    ' other processes.  Lesson learned.
    c_FileHandle = CreateFileMapping(-1&, ByVal 0&, PAGE_READWRITE, 0&, 12&, "lvIPerPropBrowsingMemMap" & App.ThreadID)
    If c_FileHandle = 0& Then ' can't create mapped file; abort
        bError = True
    Else
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then               ' already exists, see if new vTable is being hacked
            hView = MapViewOfFile(c_FileHandle, SECTION_MAP_WRITE, 0&, 0&, 0&) ' access mapped file
            CopyMemory arrayPtr, ByVal hView + c_UncompiledOCX, 4&      ' get array pointer from mapped file
            If arrayPtr = 0& Then
                bNew = True
            Else
                CopyMemory vCount, ByVal arrayPtr, 4&                       ' get the count
                ' get the VTableArrayStruct data & see if the vTable already subclassed
                OverlayArrayOnMemory tSA, VarPtrArray(theTables), arrayPtr + 4&, 28&, vCount ' overlay our array on memory array
                Index = FindVTable(theTables(), VTable)                     ' search for the vTable
                If Index = vCount Then                                      ' not found, need to increment our VTable array
                    bNew = True
                Else                                        ' found and need to add new client
                    theClient = theTables(Index).ClientPtr  ' get pointer to VTable's clients
                    CopyMemory cCount, ByVal theClient, 4&  ' get the client count
                    cCount = cCount + 1&                    ' increment count & redim preserve the array
                    theClient = CoTaskMemRealloc(theClient, cCount * 12& + 4&)
                    If theClient = 0& Then
                        bError = True                       ' out of memory? can't redim the aray
                    Else                                    ' update this vTable's client pointer & its client pointer in the thunk
                        theTables(Index).ClientPtr = theClient
                        CopyMemory ByVal theTables(Index).ThunkPtr + ThunkVTableOffset, theClient + 4&, 4&
                    End If
                End If
                OverlayArrayOnMemory tSA, VarPtrArray(theTables), 0&, 0&, 0& ' remove overlay
            End If
        Else
            hView = MapViewOfFile(c_FileHandle, SECTION_MAP_WRITE, 0&, 0&, 0&) ' access mapped file
            CopyMemory ByVal hView, vTableData, 12&     ' ensure 12 bytes are zeroed out
            bNew = True                                 ' need a new vTable array
        End If
        
        If Not bError Then
            If bNew Then                                ' VTable is not in our array yet
                theClient = CoTaskMemAlloc(16&)         ' create client array, containing single client (16 bytes)
                If theClient = 0& Then                  ' check for error
                    bError = True
                Else
                    cCount = 1&                         ' initialize count
                    vCount = vCount + 1                 ' increment vtable array count & ReDim Preserve array
                    arrayPtr = CoTaskMemRealloc(arrayPtr, vCount * 28& + 4&)
                    If arrayPtr = 0& Then
                        bError = True                   ' out of memory? couldn't redim the array
                    Else
                        OverlayArrayOnMemory tSA, VarPtrArray(theTables), arrayPtr + 4&, 28&, vCount  ' overlay our array on memory array
                        theTables(vCount - 1).ThunkPtr = CreateThunk(theClient + 4&, True) ' create thunk for this vTable
                        If theTables(vCount - 1).ThunkPtr Then                          ' if success then....
                            With theTables(vCount - 1)                                  ' update the data
                                .VTable = VTable                                        ' ref the vTable; needed to restore pointers
                                .ClientPtr = theClient                                  ' ref the client array, used by GetClientPtr
                                CopyMemory .origFunction1, ByVal VTable + 12&, 4&       ' cache original function pointers
                                CopyMemory .origFunction2, ByVal VTable + 20&, 4&
                                CopyMemory .origFunction3, ByVal VTable + 24&, 4&
                                           .origFunction4 = 0&                          ' zero out, may contain garbage. Hacked on demand only
                                SetVTableEntry VTable + 12&, .ThunkPtr + ThunkGDSOffset ' hack the vtable with thunk's functions
                                SetVTableEntry VTable + 20&, .ThunkPtr + ThunkGPDSOffset
                                SetVTableEntry VTable + 24&, .ThunkPtr + ThunkGPDVOffset
                            End With
                            CopyMemory ByVal arrayPtr, vCount, 4&                       ' add the count for the array
                            CopyMemory ByVal hView + c_UncompiledOCX, arrayPtr, 4&      ' write the array pointer to the mapped file
                        Else
                            bError = True
                        End If
                        OverlayArrayOnMemory tSA, VarPtrArray(theTables), 0&, 0&, 0&    ' remove overlay
                    End If
                End If
            End If
        End If
        UnmapViewOfFile hView
        
        If bError Then
            If theClient Then CoTaskMemFree theClient       ' free any client memory created
        Else                                                ' else add the client
            OverlayArrayOnMemory tSA, VarPtrArray(theClients), theClient + 4&, 12&, cCount
            ' patch to try to ignore issue when control is uncompiled and user changes it while
            ' it is still displayed. VB will cross-hatch the control and when user returns to
            ' the updated control or uses "Update Controls" menu item, VB will trash the old
            ' controls and recreate them; however, it doesn't call the terminate events and
            ' the pointers to the class are still in thunk but are now invalid. But VB does not
            ' create a new interface. So we simply slide the older pointers to the right and
            ' when this class' GetClientPointer is called, it will look from left to right
            ' and should always find the correct one before the bad pointers are even looked at.
            If Not bNew Then
                CopyMemory theClients(1), theClients(0), (cCount - 1&) * 12&
            End If
            With theClients(0)                              ' update the client data now
                .MePointer = ObjPtr(Me)                     ' reference to this class
                .InterfacePtr = c_InterfaceIPPB                 ' reference the usercontrols interface object
                .LocalVTableOffset = c_localVTableOffset    ' add this class' function pointers that will be called by thunk
            End With
            OverlayArrayOnMemory tSA, VarPtrArray(theClients), 0&, 0&, 0&
            CopyMemory ByVal theClient, cCount, 4&          ' update the client count
            AddClient_IPerPropertyBrowsing = True
        End If
    End If
End Function

Private Function AddClient_IPropertyPage() As Boolean

    ' This routine is very similar to the AddClient_IPerPropertyBrowsing function
    ' except that it sets up only those properties that will open a custom dialog

    Dim CLSID(0 To 3) As Long, GUID(0 To 3) As Long
    Dim Clients() As VTableArrayClientStruct, tSA As SafeArray
    Dim cCount As Long, vAddr As Long, cAddr As Long, lRtn As Long
    Dim ppgVTable As VTableArrayStruct, vTables() As VTableArrayStruct
    Dim bError As Boolean, hView As Long, ppgInterface As Long
    
    Const IID_IPropertyPage As String = "{B196B28D-BAB4-101A-B69C-00AA00341D07}"

    hView = MapViewOfFile(c_FileHandle, SECTION_MAP_WRITE, 0&, 0&, 0&)
    CopyMemory vAddr, ByVal hView + 8&, 4&                                  ' get memory address for vTable array (if available)
    If vAddr = 0& Then                                                      ' need to create the thunk and array?
        ' here is some magic & trickery.
        ' We want to be able to be notified when VB is about to show a dialog/property page when
        ' user clicks on the ellipsis buttton "..." and in order to do that, we need to know which
        ' object will be invoked when that button is clicked.  vbAccelerator did this by having VB
        ' invoke a class within its DLL, but we don't have a DLL and our class when uncompiled does
        ' not have a permanent GUID/CLSID yet. Hmmmm... So, we will actually use one of VB's standard
        ' property sheets which does have a permanent CLSID (the standard color property page).
        ' By creating an instance of this, we can attach an IPropertyPage interface to it, and get
        ' a pointer to that interface using the following function.  But how do we tell VB that
        ' the usercontrol will invoke this property page? We do that in the MapPropertyToPage function
        CLSID(0) = 2126359265: CLSID(1) = 298811680: CLSID(2) = -1442799735: CLSID(3) = 277571584
        Call ConvertGUIDtoArray(IID_IPropertyPage, GUID())
        ' create property page instance & attach IPropertyPage
        Call CoCreateInstance(CLSID(0), 0&, CLSCTX_INPROC_SERVER, GUID(0), ppgInterface)
        If ppgInterface Then
            cAddr = CoTaskMemAlloc(36&)                                     ' create initial client array (12 + extra 24)
            If cAddr Then
                vAddr = CoTaskMemAlloc(28&)                                 ' create vTable array
                If vAddr Then
                    ppgVTable.ThunkPtr = CreateThunk(cAddr + 24&, False)    ' create thunk for IPropertyPage
                    If ppgVTable.ThunkPtr Then
                        With ppgVTable                                      ' update vTable array
                            .ClientPtr = cAddr                              ' pointer to client array
                            CopyMemory .VTable, ByVal ppgInterface, 4&      ' pointer to IPropertyPage's vTable
                            CopyMemory .origFunction1, ByVal .VTable + 16, 4& ' cache pointer to IPropertyPage:Activate
                                        .origFunction2 = ppgInterface       ' cache for eventual release
                            SetVTableEntry .VTable + 16&, .ThunkPtr + ThunkActivateOffset ' hack IPropertyPage:Activate
                        End With
                        CopyMemory ByVal vAddr, ppgVTable, 28&              ' copy vTable structure to memory
                        CopyMemory ByVal hView + 8&, vAddr, 4&              ' add vTable's memory address to mapped file
                        cCount = 1&                                         ' increment client count
                        OverlayArrayOnMemory tSA, VarPtrArray(Clients), cAddr, 12&, cCount + 2  ' overlay local array on client memory
                        Clients(0).MePointer = GetProcAddress(GetModuleHandleA("user32"), "SetWindowLongA")     ' Store the SetWindowLong function address in the thunk data
                        Clients(0).InterfacePtr = GetProcAddress(GetModuleHandleA("user32"), "CallWindowProcA") ' Store CallWindowProc function address in the thunk data
                        'Clients(0).LocalVTableOffset will be set in IPropertyPage_Activate                     ' Store previous window proc
                        Clients(1).InterfacePtr = 0&: Clients(1).LocalVTableOffset = 0&              ' zero these out; may contain garbage otherwise
                        OverlayArrayOnMemory tSA, VarPtrArray(Clients), cAddr + 12&, 12&, cCount + 1 ' re-overlay local array on client memory, offsetting
                    Else
                        bError = True ' can't create thunk
                    End If
                Else
                    bError = True   ' can't create vTable array
                End If
            Else
                bError = True   ' can't create client array
            End If
        End If
    Else                                                                    ' add this class as a client?
        CopyMemory ppgVTable, ByVal vAddr, 28&                              ' get vTable structure from memmory
        CopyMemory cCount, ByVal ppgVTable.ClientPtr + 12&, 4&              ' get number of current clients
        OverlayArrayOnMemory tSA, VarPtrArray(Clients), ppgVTable.ClientPtr + 12&, 12&, cCount + 1
        lRtn = ObjPtr(Me)
        For cAddr = 1 To cCount
            If Clients(cAddr).MePointer = lRtn Then Exit For
        Next
        If cAddr > cCount Then                                              ' not a client yet
            cCount = cCount + 1&                                                ' increment count
            cAddr = CoTaskMemRealloc(ppgVTable.ClientPtr, cCount * 12& + 24&)   ' redim preserve the client array
            If cAddr Then
                CopyMemory ByVal vAddr + 4&, cAddr, 4&                          ' update client array pointer
                CopyMemory ByVal ppgVTable.ThunkPtr + ThunkVTableOffsetPPG, cAddr + 24&, 4&
                OverlayArrayOnMemory tSA, VarPtrArray(Clients), cAddr + 12&, 12&, cCount + 1
                CopyMemory Clients(2), Clients(1), (cCount - 1&) * 12&          ' shift clients right
            Else
                bError = True                                                   ' failed to resize the array
                vAddr = 0&                                                      ' prevent removing vTable array from memory
            End If
        Else
            cCount = 0&                                                         ' already a client, no updating needed
        End If
    End If
    
    If bError Then                                                          ' couldn't create needed arrays (very low memory available)
        If vAddr Then CoTaskMemFree vAddr
        If cAddr Then CoTaskMemFree cAddr
    Else
        If cCount Then
            With Clients(1)                                                     ' update new client info
                .MePointer = ObjPtr(Me)                                         ' pointer to this class
                .LocalVTableOffset = c_localVTableOffset                        ' class functions offset
            End With
            ' Clients(0) members are really: ClientCount, ActiveClientObjPtr, ActiveClientVTableOffset
            Clients(0).MePointer = cCount
            OverlayArrayOnMemory tSA, VarPtrArray(Clients), 0&, 0&, 0&          ' remove overlay
            
            ' now we are moving back to the IPPB vs IPPG to validate MapPropertyToPage is hacked
            CopyMemory vAddr, ByVal hView + c_UncompiledOCX, 4&                 ' verify MapPropertyToPage is hooked (done on demand)
            If vAddr Then
                CopyMemory cCount, ByVal vAddr, 4&                              ' need to find VTable entry for our interface
                OverlayArrayOnMemory tSA, VarPtrArray(vTables), vAddr + 4&, 28&, cCount
                CopyMemory cAddr, ByVal c_InterfaceIPPB, 4&                     ' get our interface
                cAddr = FindVTable(vTables(), cAddr)                            ' find the table
                If cAddr < cCount Then
                    If vTables(cAddr).origFunction4 = 0& Then                   ' haven't hacked it yet, do it now
                        CopyMemory vTables(cAddr).origFunction4, ByVal vTables(cAddr).VTable + 16&, 4&
                        SetVTableEntry vTables(cAddr).VTable + 16&, vTables(cAddr).ThunkPtr + ThunkMPTPOffset
                    End If
                End If
                OverlayArrayOnMemory tSA, VarPtrArray(vTables), 0&, 0&, 0&      ' remove overlay
            End If
            tSA.pvData = 0&
        End If
        AddClient_IPropertyPage = True
    End If
    UnmapViewOfFile hView
    If tSA.pvData Then OverlayArrayOnMemory tSA, VarPtrArray(Clients), 0&, 0&, 0&

End Function

Private Sub OverlayArrayOnMemory(tSA As SafeArray, arrayPtr As Long, memPtr As Long, elementSize As Long, elementCount As Long, Optional lowBound As Long = 0&)
    ' Helper function to overlay a local VB array onto some memory or remove the overlay
    If memPtr = 0& Then
        CopyMemory ByVal arrayPtr, 0&, 4&
    Else
        With tSA
            .cDims = 1
            .cbElements = elementSize
            .pvData = memPtr
            .rgSABound.cElements = elementCount
            .rgSABound.lLbound = lowBound
        End With
        CopyMemory ByVal arrayPtr, VarPtr(tSA), 4&
    End If
End Sub

Private Function CreateThunk(memOffset As Long, StandardThunk As Boolean) As Long
    ' This function creates the memory thunk that will interact between the
    ' subclassed IPerPropertyBrowsing/IPropertyPage interfaces and instances of these classes.
    
    ' The thunks are provided in the attached IPPBCallback.asm/IPPGCallback.asm & can be viewed in Notepad/WordPad
    
    Dim z_Code()        As Long
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Dim MEM_LEN         As Long                 'Length of z_Code in bytes
    Dim apiThunk        As Long                 'memory for thunk

    If StandardThunk Then
'         IPPBCallback thunk's logic basically looks like the following
'         --------------------------------------------------------------
'        Function Main(IPPB_Interface_Pointer, 2 or 3 other parameters) As Long
'           Get VTable's Client pointer (supplied as z_Code(51) below)
'           Do clients exist?
'              If Yes Then
'                Do more than one client exist?
'                    If so, call this class's IPerPropertyBrowsing_HelperFunction, passing it the interface pointer & current client pointer
'                    This class will return the correct class pointer to call to
'                Call the class function for the interface function that called the thunk
'                Return to the interface what the called class returns
'          Otherwise
'                Return E_NOTIMPL to the interface if no client info found
'        End Function

        ReDim z_Code(0 To 65)
        z_Code(0) = &HE8E58960: z_Code(1) = &HC0&: z_Code(2) = &H1D74F739: z_Code(3) = &H501C458D: z_Code(4) = &HFF2C75FF: z_Code(5) = &H75FF2875: z_Code(6) = &H8B37FF24: z_Code(7) = &HE2D10447: z_Code(8) = &HD0FFD001: z_Code(9) = &H7E9&: z_Code(10) = &H1C45C700: z_Code(11) = &H80004001: z_Code(12) = &HCC261: z_Code(13) = &HE8E58960: z_Code(14) = &H8C&: z_Code(15) = &H5174F739: z_Code(16) = &H501C458D: z_Code(17) = &HFF3075FF: z_Code(18) = &H75FF2C75: z_Code(19) = &H2475FF28: z_Code(20) = &H478B37FF: z_Code(21) = &HD1D00104: z_Code(22) = &HFFD001E2: z_Code(23) = &H36E9D0: z_Code(24) = &H90900000: z_Code(25) = &HE8E58960: z_Code(26) = &H5C&: z_Code(27) = &H2174F739: z_Code(28) = &H501C458D: z_Code(29) = &HFF3075FF: z_Code(30) = &H75FF2C75: z_Code(31) = &H2475FF28: z_Code(32) = &H478B37FF: z_Code(33) = &H2E2C104: z_Code(34) = &HD0FFD001: z_Code(35) = &H7E9&: z_Code(36) = &H1C45C700
        z_Code(37) = &H80004001: z_Code(38) = &H10C261: z_Code(39) = &HE8E58960: z_Code(40) = &H24&: z_Code(41) = &H8174F739: z_Code(42) = &H501C458D: z_Code(43) = &HFF2C75FF: z_Code(44) = &H75FF2875: z_Code(45) = &H8B37FF24: z_Code(46) = &HD0010447: z_Code(47) = &H102E2C1: z_Code(48) = &HE9D0FFD0: z_Code(49) = &HFFFFFF68: z_Code(50) = &HBFF63190: z_Code(51) = &H12345678: z_Code(52) = &HF739C031: z_Code(53) = &H478B2F74: z_Code(54) = &H75F039FC: z_Code(55) = &HC3FF3103: z_Code(56) = &HBA909090: z_Code(57) = &H12345678: z_Code(58) = &H13D&: z_Code(59) = &H52167400: z_Code(60) = &H501C458D: z_Code(61) = &H2475FF57: z_Code(62) = &H478B37FF: z_Code(63) = &HFFD00104: z_Code(64) = &H1C7D8BD0: z_Code(65) = &HC35A&
        z_Code(57) = zAddressOf(Me, 5) - c_localVTableOffset ' compiled/uncompiled classes have different vTable offsets between contigouos functions
        z_Code(51) = memOffset
        ' module variable: ThunkVTableOffset must = above z_Code index * 4 (i.e., 57*4)
        
    Else
'         IPPGCallback thunk's logic basically looks like the following
'         --------------------------------------------------------------
'        Function Main(IPPB_Interface_Pointer, 3 other parameters) As Long
'           Get VTable's Client pointer (supplied as z_Code(39) below)
'           Do clients exist?
'               If Yes Then
'                   Move back 8 bytes from client pointer & check for active client information
'                   Does active client exist?
'                   If No Then call any client to forward the hacked event to calling pointer
'                   If Yes then call this class' IPropertyPage_Activate event
'          Otherwise
'               Return E_UNEXPECTED to the interface if no client info found
'       End Function

        ReDim z_Code(0 To 47)
        z_Code(0) = &HE8E58960: z_Code(1) = &H90&: z_Code(2) = &H82287D81: z_Code(3) = &H74000000: z_Code(4) = &H2FE80A: z_Code(5) = &H76E90000: z_Code(6) = &H8D000000: z_Code(7) = &H56501C45: z_Code(8) = &HFF565656: z_Code(9) = &H457FF37: z_Code(10) = &H18E8&: z_Code(11) = &HF477FF00: z_Code(12) = &HFFFFFC68: z_Code(13) = &H2475FFFF: z_Code(14) = &H10EF81: z_Code(15) = &H17FF0000: z_Code(16) = &H4CE9&: z_Code(17) = &H3075FF00: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875: z_Code(20) = &HF477FF24: z_Code(21) = &H82DF889: z_Code(22) = &HFF000000: z_Code(23) = &H1C458910: z_Code(24) = &H909090C3: z_Code(25) = &HE8E58960: z_Code(26) = &H2C&: z_Code(27) = &H1A74F739: z_Code(28) = &H501C458D: z_Code(29) = &HFF3075FF: z_Code(30) = &H75FF2C75: z_Code(31) = &H2475FF28: z_Code(32) = &H57FF37FF: z_Code(33) = &H7E904: z_Code(34) = &H45C70000: z_Code(35) = &HFFFF1C: z_Code(36) = &H10C26180
        z_Code(37) = &H90909000: z_Code(38) = &HBFF63190: z_Code(39) = &H12345678: z_Code(40) = &HF739C031: z_Code(41) = &H478B1774: z_Code(42) = &H75F039F4: z_Code(43) = &HC3FF3103: z_Code(44) = &H39F8478B: z_Code(45) = &H810674F0: z_Code(46) = &H8EF&: z_Code(47) = &HC300&
        z_Code(39) = memOffset
        ' module variable: ThunkVTableOffsetPPG must = above z_Code index * 4 (i.e., 39*4)

    End If
    
    MEM_LEN = UBound(z_Code) * 4& + 4&
    apiThunk = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)   'Allocate executable memory
    If apiThunk = 0& Then Exit Function                         'Ensure the allocation succeeded
    
    CopyMemory ByVal apiThunk, z_Code(0), MEM_LEN               'Copy the thunk code/data to the allocated memory
    CreateThunk = apiThunk                                      'Return pointer
    
End Function

Private Function FindVTable(vTables() As VTableArrayStruct, vTableRef As Long) As Long
    ' Helper function to iterate thru VTable array
    Dim Index As Long
    For Index = 0& To UBound(vTables)
        If vTables(Index).VTable = vTableRef Then Exit For
    Next
    FindVTable = Index
End Function

Private Function FindClient(Clients() As VTableArrayClientStruct, IFaceRef As Long) As Long
    ' Helper function to iterate thru VTable's Client array
    Dim Index As Long
    For Index = 0& To UBound(Clients)
        If Clients(Index).InterfacePtr = IFaceRef Then Exit For
    Next
    FindClient = Index
End Function

Private Function ConvertGUIDtoArray(ByVal sGUID As String, aGUID() As Long) As Boolean
    ' Helper function to convert string GUID to 16 bytes
    On Error GoTo CATCH_EXCEPTION
    sGUID = Mid$(Replace$(sGUID, "-", vbNullString), 2, 32)
    If Len(sGUID) = 32 Then
        aGUID(0) = CLng("&H" & Left$(sGUID, 8))
        aGUID(1) = CLng("&H" & Mid$(sGUID, 13, 4) & Mid$(sGUID, 9, 4))
        aGUID(2) = CLng("&H" & Mid$(sGUID, 23, 2) & Mid$(sGUID, 21, 2) & _
                             Mid$(sGUID, 19, 2) & Mid$(sGUID, 17, 2))
        aGUID(3) = CLng("&H" & Mid$(sGUID, 31, 2) & Mid$(sGUID, 29, 2) & _
                             Mid$(sGUID, 27, 2) & Mid$(sGUID, 25, 2))
        ConvertGUIDtoArray = True
    End If
CATCH_EXCEPTION:
    If Err Then Err.Clear 'passed invalid GUID format
End Function

Private Sub SetVTableEntry(OldPtr As Long, newPtr As Long)
    ' Helper function to set/replace VTable entries, called by both AddClient_IPerPropertyBrowsing & Detach
    Dim vProtect As Long
    If Not (newPtr = 0& Or OldPtr = 0&) Then
        VirtualProtect OldPtr, 4&, PAGE_EXECUTE_READWRITE, vProtect
        CopyMemory ByVal OldPtr, newPtr, 4&
        VirtualAlloc OldPtr, 4&, vProtect, vProtect
    End If
End Sub

Private Function TestArray(ByVal notArrayName As Long) As Long
    ' Helper function to test array as uninitialized or not
    TestArray = Not notArrayName
    Debug.Assert App.hInstance ' IDE bug with using Not Not Pointers. Prevents bug.
End Function

Private Function FindDispID(ID As Long, AddNew As Boolean) As Long

    ' BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Dispatch IDs for implemented properties are kept in a sorted list
    ' to aid in finding them faster.  This prevents passing non-implementted
    ' properties to the user and having them decide whether or not it is
    ' one of the ones that they have implemented.
    
    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If c_DispIDCol.Count = 0& Then
        If AddNew Then
            ReDim c_DispIDCol.Params(1 To 3)
            c_DispIDCol.Count = 1&
            c_DispIDCol.Params(c_DispIDCol.Count).DispID = ID
            newIndex = c_DispIDCol.Count
        End If
    Else
        UB = c_DispIDCol.Count
        LB = 1&
        Do Until LB > UB
            newIndex = LB + ((UB - LB) \ 2&)
            If c_DispIDCol.Params(newIndex).DispID = ID Then
                Exit Do
            ElseIf c_DispIDCol.Params(newIndex).DispID > ID Then
                UB = newIndex - 1&
            Else ' new color is higher in sort order
                LB = newIndex + 1&
            End If
        Loop
        If LB > UB Then  ' color was not found
            If AddNew Then
                If ID > c_DispIDCol.Params(newIndex).DispID Then newIndex = newIndex + 1&
                If c_DispIDCol.Count = UBound(c_DispIDCol.Params) Then
                    ReDim Preserve c_DispIDCol.Params(1 To c_DispIDCol.Count + 3&)
                End If
                With c_DispIDCol
                    .Count = .Count + 1&
                    If newIndex < .Count Then
                        For LB = .Count To newIndex + 1& Step -1&
                            .Params(LB) = .Params(LB - 1&)
                        Next
                    End If
                    .Params(newIndex).DispID = ID
                End With
            Else
                newIndex = 0&
            End If
        End If
    End If
    FindDispID = newIndex
End Function

Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
' Code borrowed from one of Paul Caton's excellent thunk projects
  
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  CopyMemory nAddr, ByVal ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                               'Probe for a Class method
    If Not zProbe(nAddr + &H7A4, i, bSub) Then                                'Probe for a UserControl method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                          'Probe for a Form method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    CopyMemory nAddr, ByVal i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    CopyMemory bVal, ByVal nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4              'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  ' Helper function for zAddressOf
  ' Probes at the specified start address for a method signature
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function IsUnCompiled() As Boolean
    ' helper function to determine if this class is compiled or not
    ' It is not called when compiled; c_UncompiledOCX=0& in that case
    c_UncompiledOCX = 4&    ' if we got here, we are uncompiled
    IsUnCompiled = True
End Function

Private Sub Class_Initialize()
    ' reference point from which all other subclassed functions are based on
    c_localVTableOffset = zAddressOf(Me, 6) ' pointer to our IPropertyPage:Activate function
    ' functions afterwards are +32/13 bytes each & offsets used in the IPerPropertyBrowsing thunk
    ' functions before are -32/13 bytes each & offsets used in the IPropertyPage thunk
    ' offsets are based on class being compiled or not.
End Sub

Private Sub Class_Terminate()
    Me.Detach   ' call clean up if not called by usercontrol
End Sub


'|------------------------------------------------------------------------------------------------------------
' //// IMPORTANT: THESE LAST 6 FUNCTIONS MUST REMAIN TOGETHER, REMAIN LAST, AND REMAIN IN THIS SPECIFIC ORDER
'|------------------------------------------------------------------------------------------------------------
'|- IPropertyPage_Activate
'|- IPerPropertyBrowsing_HelperFunction
'|- IPerPropertyBrowsing_GetDisplayString
'|- IPerPropertyBrowsing_GetPredefinedStrings
'|- IPerPropertyBrowsing_GetPredefinedValue
'|- IPerPropertyBrowsing_MapPropertyToPage
'|- absolutely no additional executable code can exist beyond IPerPropertyBrowsing_MapPropertyToPage
'|------------------------------------------------------------------------------------------------------------

Private Function IPropertyPage_Activate(ByVal This As Long, ByVal hwndParent As Long, ByVal sizeRect As Long, ByVal isModal As Long) As Long

    ' http://msdn.microsoft.com/en-us/library/ms682250(VS.85).aspx
    ' implementation of IDispatch::Activate for IPropertyPage
    
    ' If user wants to display their own dialog (say FileOpen) when the ellipsis (...) button is clicked, the
    ' IPerPropertyBrowsing_MapToPropertyPage function tricked VB into believing it is a mapped property. Before
    ' displaying the property page, VB will fire off a series of IPropertyPage events: SetPageSite, SetObjects,
    ' Activate, and more.  We can intercept the property page at any time, however, intercepting it before Activate
    ' triggers a modal messagebox displaying an error code. Even though we could possibly set up a hook and and
    ' destroy the msgbox before it is displayed, you still get beeps. So, we trap the Activate event which we can
    ' abort at also. This produces no beeps, however, you may or may not get a quick flash in IDE when the property page
    ' is destroyed. This is a bit more difficult to get exactly right. VB tends to override everything I've tried and
    ' this is the closest I've gotten to be perfect. The quick flash is intermittent if it occurs. We will actually subclass
    ' the property page while it is still invisible, so we know for sure exactly when it is destroyed.
    
    ' Note: The parameter This is another instance of IPropertyPage, but is not our instance.
    ' The other parameters refer to the property page used to host controls (i.e., buttons, textboxes, etc)
    ' If the parameters are all zeroes, then this is a callback from the subclassing thunk & not VB.
    
    Dim VTable As VTableArrayStruct
    Dim hView As Long, memOffset As Long
    Dim varRtn As Variant
    Dim pvTypes(0 To 2) As Integer, pvPtrs(0 To 2) As Long, pValues(0 To 2) As Variant
    Dim iInterface As IPropertyBrowserEvents, tObj As Object, sPropName As String

    Const E_UNEXPECTED As Long = &H8000FFFF
    Const GWL_WNDPROC As Long = &HFFFC
    
    If This = 0& Then                                                           ' callback from the subclassing thunk
        SendMessage GetDesktopWindow, WM_SETREDRAW, 1&, ByVal 0&                ' allow desktop to repaint
    Else
        hView = MapViewOfFile(c_FileHandle, SECTION_MAP_READ, 0&, 0&, 0&)       ' get mapped file pointer
        If hView = 0& Then                                                      ' error
            IPropertyPage_Activate = E_UNEXPECTED
        Else
            CopyMemory memOffset, ByVal hView + 8&, 4&                          ' get pointer to vTable structure
            UnmapViewOfFile hView                                               ' unmap file, no longer needed here
            CopyMemory VTable, ByVal memOffset, 28&                             ' copy structure
            If c_MappedDispID = 0& Then
                ' if not expecting this, then another property in this usercontrol or another usercontrol
                ' is also using the same property page that we are subclassing; therefore, we will get this
                ' event even though we don't want it...
                SetVTableEntry VTable.VTable + 16&, VTable.origFunction1        ' restore vTable's original function pointer & call that function
                pvTypes(0) = vbLong: pValues(0) = hwndParent:   pvPtrs(0) = VarPtr(pValues(0))
                pvTypes(1) = vbLong: pValues(1) = sizeRect:     pvPtrs(1) = VarPtr(pValues(1))
                pvTypes(2) = vbLong: pValues(2) = isModal:      pvPtrs(2) = VarPtr(pValues(2))
                ' HRESULT Activate(HWND hWndParent, LPCRECT prc, BOOL bModal);
                DispCallFunc This, 16&, CC_STDCALL, vbLong, 3&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn
                SetVTableEntry VTable.VTable + 16&, VTable.ThunkPtr + ThunkActivateOffset ' re-hack function pointer
                IPropertyPage_Activate = CLng(varRtn)
            Else
                ' we are expecting this
                CopyMemory tObj, c_UCIPEPointer, 4&                             ' get usercontrol's events pointer
                Set iInterface = tObj
                CopyMemory tObj, 0&, 4&
                sPropName = c_DispIDCol.Params(c_MappedDispID).Name             ' prep property name
                ' See IPropertyBrowsingEvents.ShowCustomPropertyPage for more info
                Call iInterface.ShowCustomPropertyPage(sPropName)               ' send usercontrol the event
                
                ' subclass the property page's parent/owner
                memOffset = SetWindowLongA(GetParent(hwndParent), GWL_WNDPROC, VTable.ThunkPtr)
                CopyMemory ByVal VTable.ClientPtr + 4&, memOffset, 4&           ' store where thunk expects it
                SendMessage GetDesktopWindow, WM_SETREDRAW, 0&, ByVal 0&        ' prevent desktop from repainting
                ' ^^ the above is done so the destroyed property page does not "ghost" on the IDE hDC.
                ' The subclassing will call back when the page is destroyed & we will re-enable desktop painting
                
                IPropertyPage_Activate = E_UNEXPECTED                           ' now tell parent oops
            End If
        End If
    End If
    
End Function

Private Function IPerPropertyBrowsing_HelperFunction(ByVal This As Long, ByVal memOffset As Long) As Long

    ' This is a thunk helper function. Instead of adding the necessary code in the
    ' thunk, the thunk calls this routine to locate the Class pointer for the
    ' interface it should call. Called only when more than one usercontrol shares
    ' the same IPerPropertyBrowsing interface vTable.
    
    ' Thunk passes this class the interface pointer (This) and the memory location
    ' for the 1st client of that interface (memOffset)

    Dim Count As Long
    Dim Clients() As VTableArrayClientStruct
    Dim tSA As SafeArray
    
    On Error GoTo ExitRoutine
    ' get the table's client count, 4 bytes in front of the client array address
    If memOffset Then
        CopyMemory Count, ByVal memOffset - 4&, 4&
        ' find the requested interface pointer (This)
        If Count Then
            OverlayArrayOnMemory tSA, VarPtrArray(Clients), memOffset, 12&, Count
            memOffset = FindClient(Clients(), This)
            ' if found, return the memory address of the client that the thunk needs to call
            If memOffset < Count Then IPerPropertyBrowsing_HelperFunction = VarPtr(Clients(memOffset).MePointer)
        End If
    End If
ExitRoutine:
If tSA.pvData Then OverlayArrayOnMemory tSA, VarPtrArray(Clients), 0&, 0&, 0&
End Function

Private Function IPerPropertyBrowsing_GetDisplayString(ByVal This As Long, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long

    ' http://msdn.microsoft.com/en-us/library/ms688734(VS.85).aspx
    ' implementation of IDispatch:GetDisplayString for IPerPropertyBrowsing
    
    ' This function displays in the propety sheet whatever you want displayed to the user as the property's value
    
    ' See IPropertyBrowserEvents.FormatPropertyDisplay for more info

    Dim bUseDefault As Boolean
    Dim strDisplayName As String
    Dim strPropName As String
    Dim rtnStrPtr As Long
    Dim iInterface As IPropertyBrowserEvents, tObject As Object
    
    On Error GoTo CATCH_EXCEPTION
    If VarPtr(lpDisplayName) = 0& Then
        IPerPropertyBrowsing_GetDisplayString = E_POINTER
    Else
        'handle all errors so we don't crash caller
        DispID = FindDispID(DispID, False)
        If DispID = 0& Then
            IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
        ElseIf c_DispIDCol.Params(DispID).pType = ptype_Dialog Then  ' does not apply to non-display-locking dialog
            IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
        Else
            CopyMemory tObject, c_UCIPEPointer, 4&
            Set iInterface = tObject
            CopyMemory tObject, 0&, 4&
            ' value is about to be displayed on property sheet, ask client to format it
            strPropName = c_DispIDCol.Params(DispID).Name
            Call iInterface.FormatPropertyDisplay(strPropName, strDisplayName, bUseDefault)
            
            If bUseDefault Then
                IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
            Else
                'copy display string to passed ptr (caller, which is VB, frees the memory)
                ' Per MSDN, the string must be created using SysAllocString
                If StrPtr(strDisplayName) = 0& Then strDisplayName = ""
                rtnStrPtr = SysAllocString(StrPtr(strDisplayName))
                CopyMemory ByVal lpDisplayName, rtnStrPtr, 4&
            End If
        End If
    End If
    Exit Function
    
CATCH_EXCEPTION:
    IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
    Debug.Print Err.Description, "gds"
End Function

Private Function IPerPropertyBrowsing_GetPredefinedStrings(ByVal This As Long, ByVal DispID As Long, _
                                    ByRef pCaStringsOut As CAWORDOLESTR, ByRef pCaCookiesOut As CAWORDOLESTR) As Long
    '  http://msdn.microsoft.com/en-us/library/ms679724(VS.85).aspx
    ' implementation of IDispatch:GetPredefinedStrings for IPerPropertyBrowsing
    
    ' This function allows you to format the dropdown box of enumeration values to be displayed in any string format you wish
    
    ' See IPropertyBrowserEvents.FormatPropertyEnum for more info


    Dim strPropName As String
    Dim strStringsOut() As String               ' array to pass to client to be filled with enumeration values
    Dim lngCookiesOut() As Variant              ' array to pass to client to be filled with enumeration display strings
    Dim rtnStringPtrs() As Long                 ' direct memory access overlay array
    Dim rtnCookies() As Long                    ' direct memory access overlay array
    Dim sSA As SafeArray, cSA  As SafeArray     ' safe arays for DMA overlays
    Dim sPtr As Long, cPtr As Long              ' pointers for DMA overlays
    Dim Index As Long, arraySize As Long, sLen As Long
    Dim iInterface As IPropertyBrowserEvents, tObject As Object

    On Error GoTo CATCH_EXCEPTION
    'validate passed pointers
    If VarPtr(pCaStringsOut) = 0& Or VarPtr(pCaCookiesOut) = 0& Then
        IPerPropertyBrowsing_GetPredefinedStrings = E_POINTER
    Else
        ' are we managing the passed dispatch id?
        DispID = FindDispID(DispID, False)
        If DispID = 0& Then
            IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
        ElseIf c_DispIDCol.Params(DispID).pType = ptype_Enum Then ' only applies to enum-type properties
            CopyMemory tObject, c_UCIPEPointer, 4&
            Set iInterface = tObject
            CopyMemory tObject, 0&, 4&
            ' ask client to populate arrays: arrays must be same lbound and ubound
            strPropName = c_DispIDCol.Params(DispID).Name
            Call iInterface.FormatPropertyEnum(strPropName, strStringsOut(), lngCookiesOut())
            
            If TestArray(Not strStringsOut) = 0& Then               ' user passed null array; use vb default
                IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
            Else
                If TestArray(Not lngCookiesOut) = 0& Then           ' user passed null array
                    ReDim lngCookiesOut(LBound(strStringsOut) To UBound(strStringsOut))
                    For Index = LBound(strStringsOut) To UBound(strStringsOut)
                        lngCookiesOut(Index) = Index
                    Next
                Else ' validate arrays have same LBound/UBound
                    If LBound(lngCookiesOut) <> LBound(strStringsOut) Then ' must be same lbound (local requirement); use vb default
                        IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
                        Exit Function
                    End If
                    If UBound(lngCookiesOut) <> UBound(strStringsOut) Then ' must be same ubound (local requirement)
                        IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
                        Exit Function
                    End If
                End If
                arraySize = UBound(strStringsOut) + 1&                  ' how many array items?
                ' per MSDN, the arrays & strings must be created using CoTaskMemAlloc
                '           the caller, which is VB, is responsible for freeing memory
                cPtr = CoTaskMemAlloc(arraySize * 4&)                   ' create memory for return array
                If cPtr = 0& Then
                    IPerPropertyBrowsing_GetPredefinedStrings = E_OUTOFMEMORY                ' abort
                Else
                    sPtr = CoTaskMemAlloc(arraySize * 4&)               ' create memory for return array
                    If sPtr = 0& Then
                        CoTaskMemFree cPtr                              ' abort
                        IPerPropertyBrowsing_GetPredefinedStrings = E_OUTOFMEMORY
                    Else
                        ' overlay local arrays on the memory addresses to make looping easier
                        OverlayArrayOnMemory cSA, VarPtrArray(rtnCookies), cPtr, 4&, arraySize, LBound(strStringsOut)
                        OverlayArrayOnMemory sSA, VarPtrArray(rtnStringPtrs), sPtr, 4&, arraySize, LBound(strStringsOut)
                        ' loop thru each client-returned array & set to the memory arrays
                        For Index = LBound(strStringsOut) To UBound(strStringsOut)
                            sLen = LenB(strStringsOut(Index)) + 2&              ' get len of string & add null terminator
                            rtnStringPtrs(Index) = CoTaskMemAlloc(sLen)         ' create space for the string & copy the string data
                            CopyMemory ByVal rtnStringPtrs(Index), ByVal StrPtr(strStringsOut(Index)), sLen
                            rtnCookies(Index) = CLng(Val(lngCookiesOut(Index))) ' copy the itemdata for this string item
                        Next
                        ' remove overlays; no longer needed
                        OverlayArrayOnMemory cSA, VarPtrArray(rtnCookies), 0&, 0&, 0&
                        OverlayArrayOnMemory sSA, VarPtrArray(rtnStringPtrs), 0&, 0&, 0&
                        ' update the UDT members & copy to parameters
                        pCaCookiesOut.Count = arraySize
                        pCaStringsOut.Count = arraySize
                        pCaCookiesOut.DataPtr = cPtr
                        pCaStringsOut.DataPtr = sPtr
                    End If
                End If
            End If
        Else
            IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
        End If
    End If
    Exit Function
    
CATCH_EXCEPTION:
    If cPtr Then CoTaskMemFree cPtr
    If sPtr Then CoTaskMemFree sPtr
    IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
    Debug.Print Err.Description, "gps"
End Function

Private Function IPerPropertyBrowsing_GetPredefinedValue(ByVal This As Long, ByVal DispID As Long, ByVal ItemData As Long, ByRef varOut As Variant) As Long
    ' http://msdn.microsoft.com/en-us/library/ms690401(VS.85).aspx
    ' implementation of IDispatch:GetPredefinedValue for IPerPropertyBrowsing
    
    ' Per MSDN documentation, this must return E_NOTIMPL if the property did not implement GetPredefinedStrings
    ' Therefore, when the AddProperty function was called by user and DisplayDropDownSelection was False, this function is not implemented
    
    ' See IPropertyBrowserEvents.SetEnumPropertyValue for more info
    
    Dim sPropName As String
    Dim iInterface As IPropertyBrowserEvents, tObject As Object
    
    On Error GoTo CATCH_EXCEPTION
    'validate passed pointers
    If VarPtr(ItemData) = 0 Or VarPtr(varOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedValue = E_POINTER
    Else
        ' are we managing the passed dispatch id?
        DispID = FindDispID(DispID, False)
        If DispID = 0& Then
            IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        ElseIf c_DispIDCol.Params(DispID).pType = ptype_Enum Then  ' only applies to enum-type properties
            CopyMemory tObject, c_UCIPEPointer, 4&
            Set iInterface = tObject
            CopyMemory tObject, 0&, 4&
            ' ask client to set the property's value
            sPropName = c_DispIDCol.Params(DispID).Name
            Call iInterface.SetEnumPropertyValue(sPropName, ItemData, varOut)
        Else
            IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        End If
    End If
    Exit Function
    
CATCH_EXCEPTION:
    IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
    Debug.Print Err.Description, "gpv"
End Function

Private Function IPerPropertyBrowsing_MapPropertyToPage(ByVal This As Long, ByVal DispID As Long, ByRef lCLSID As Long) As Long
    
    ' http://msdn.microsoft.com/en-us/library/ms694476(VS.85).aspx
    ' implementation of IDispatch:MapPropertyToPage for IPerPropertyBrowsing
    
    Dim CLSID(0 To 3) As Long
    Dim vTables() As VTableArrayStruct, tSA As SafeArray
    Dim Count As Long, Index As Long, hView As Long, VTable As Long
    Dim varRtn As Variant
    Dim pvTypes(0 To 1) As Integer, pvPtrs(0 To 1) As Long, pValues(0 To 1) As Variant
    
    c_MappedDispID = 0&                             ' reset to zero
    Index = FindDispID(DispID, False)               ' see if this dispID is redirected
    hView = MapViewOfFile(c_FileHandle, SECTION_MAP_READ, 0&, 0&, 0&)
    If hView Then
        If Index Then                               ' does not apply to enum-type properties
            If c_DispIDCol.Params(Index).pType = ptype_Enum Then Index = 0&
        End If
        If Index = 0& Then                            ' not user-redirected
            CopyMemory VTable, ByVal hView + c_UncompiledOCX, 4& ' get pointer to vTable array
            CopyMemory Count, ByVal VTable, 4&      ' get number of vTables & overlay local array
            OverlayArrayOnMemory tSA, VarPtrArray(vTables), VTable + 4&, 28&, Count
            CopyMemory VTable, ByVal This, 4&       ' get the interfaces vtable
            Index = FindVTable(vTables(), VTable)   ' see if we are managing this; should be
            If Index < Count Then                   ' replace vTable function pointer & call original function
                SetVTableEntry VTable + 16&, vTables(Index).origFunction4
                pvTypes(0) = vbLong: pValues(0) = DispID: pvPtrs(0) = VarPtr(pValues(0))
                pvTypes(1) = vbLong: pValues(1) = VarPtr(lCLSID): pvPtrs(1) = VarPtr(pValues(1))
                ' See if the usercontrol implements IPerPropertyBrowsing & it should if it has any public variables or methods
                Call DispCallFunc(This, 16&, CC_STDCALL, vbLong, 2&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
                SetVTableEntry VTable + 16&, vTables(Index).ThunkPtr + ThunkMPTPOffset ' rehack function pointer
                IPerPropertyBrowsing_MapPropertyToPage = CLng(varRtn)
            Else                                    ' should never get here
                IPerPropertyBrowsing_MapPropertyToPage = E_NOINTERFACE
            End If
            OverlayArrayOnMemory tSA, VarPtrArray(vTables), 0&, 0&, 0& ' remove overlay
            CopyMemory VTable, ByVal hView + 8&, 4& ' get property page vTable structure pointer
            CopyMemory VTable, ByVal VTable + 4, 4& ' get client pointer
            CopyMemory ByVal VTable + 16&, 0&, 4&   ' set active client to zero
        Else
            CLSID(0) = 2126359265: CLSID(1) = 298811680: CLSID(2) = -1442799735: CLSID(3) = 277571584
            CopyMemory lCLSID, CLSID(0), 16&        ' set map to managed property page
            c_MappedDispID = Index                  ' set local variable to active property
            CopyMemory VTable, ByVal hView + 8&, 4& ' get property page vTable pointer
            CopyMemory VTable, ByVal VTable + 4, 4& ' get client pointer
            CopyMemory ByVal VTable + 16&, ObjPtr(Me), 4&            ' set active client
            CopyMemory ByVal VTable + 20&, c_localVTableOffset, 4&   ' set active client's function pointer offset
        End If
        UnmapViewOfFile hView
    Else
        IPerPropertyBrowsing_MapPropertyToPage = E_NOINTERFACE
    End If
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' IMPORTANT:  DO NOT PLACE ANY EXECUTABLE CODE BELOW THIS LINE. THUNKS WILL FAIL


