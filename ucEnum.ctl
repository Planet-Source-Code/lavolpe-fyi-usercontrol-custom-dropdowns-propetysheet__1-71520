VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ucEnum 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   ScaleHeight     =   3030
   ScaleWidth      =   7635
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Notice the FileName property?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   1635
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   $"ucEnum.ctx":0000
      ForeColor       =   &H00000000&
      Height          =   930
      Index           =   7
      Left            =   4065
      TabIndex        =   7
      Top             =   1980
      Width           =   3150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Notice the ImageList property?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   6
      Left            =   4410
      TabIndex        =   6
      Top             =   330
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "It uses an on-the-fly, string enumeration  to populate the dropdown list with actual hosted ImageList controls."
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   5
      Left            =   4110
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "It uses an on-the-fly, string enumeration  to populate the dropdown list."
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   4
      Left            =   435
      TabIndex        =   4
      Top             =   1980
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "One is standard VB enumeration while the other is a custom displayed list of the same VB enumeration"
      ForeColor       =   &H00000000&
      Height          =   705
      Index           =   1
      Left            =   510
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Notice the Drives property?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   750
      TabIndex        =   3
      Top             =   1635
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Notice the TextAlign properties?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   825
      TabIndex        =   2
      Top             =   330
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Click Me Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   2715
      TabIndex        =   0
      Top             =   30
      Width           =   2355
   End
End
Attribute VB_Name = "ucEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
                                (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Note that the CommonDialog control is provided for ease only.
' If this were a real usercontrol, you might opt for a class version of the control
' to allow unicode capabiliity, callbacks, and other functionality not provided via the control

' Additionally, to provide a really professional appearance, you should probably create
' some hook code so you can display your dialogs exactly where you want too (i.e., middle
' of screen maybe).  Leave that up to you & code is available on the net for that task.

' NICE TIP      NICE TIP      NICE TIP      NICE TIP      NICE TIP      NICE TIP
' ------------------------------------------------------------------------------
' See the InUserMode function at bottom of page. This might apply to your uc too.


'/////// sample project variables only
Dim m_Drives() As String        ' used for dynamic enumeration (the Drives property)
Dim m_IndexDrive As Long        ' index of the selected m_Drives() item
Dim m_ImageLists() As String    ' used for dynamic enumeration (the ImageList property)
Dim m_IndexImgLst As Long       ' index of the selected m_ImageLists() item
Dim m_Drive As String                ' actual property value
Dim m_ImageList As Control           ' control reference
Dim m_sImageList As String           ' actual property value
Dim m_Align1 As AlignmentConstants   ' actual property value
Dim m_Align2 As AlignmentConstants   ' actual property value
Dim m_FileName As String             ' actual property value

'/////// required declarations if implementation is used
Dim c_CustomProps As cCustomPropertyDisplay
Implements IPropertyBrowserEvents

'/////// Custom dialog example - file dialog
Public Property Get FileName() As String
    ' here we are formatting the display because we
    ' did not opt to receive FormatPropertyDisplay events
    If InUserMode Then
        FileName = m_FileName
    ElseIf m_FileName = "" Then
        FileName = "{No File Selected}"
    Else
        FileName = "{File}"
    End If
End Property
Public Property Let FileName(ByVal theFile As String)
    If InUserMode Then
        ' you probably want to validate this during runtime
    End If
    m_FileName = theFile
End Property

'/////// Custom dialog example - color picker
'        Don't use OLE_COLOR else VB will use its color picker
' here we are NOT formatting the display because we opted to
' receive FormatPropertyDisplay events and the formatting
' occurs in that event, not this event
Public Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewColor As Long)
    ' probably want to verify valid color
    On Error Resume Next
    UserControl.BackColor = NewColor
    If Err Then
        ' raise error?
        Err.Clear
    Else
        PropertyChanged "BackColor"
    End If
End Property

'/////// Non-implemented VB enumeration
Public Property Get TextAlign1() As AlignmentConstants
    TextAlign1 = m_Align1
End Property
Public Property Let TextAlign1(ByVal Align As AlignmentConstants)
    ' should validate passed value is within range
    If Align < vbLeftJustify Or Align > vbCenter Then
        ' should generate invalid property value error here
        Exit Property
    End If
    m_Align1 = Align
    PropertyChanged "Align1"
End Property

'/////// Implemented VB enumeration
Public Property Get TextAlign2() As AlignmentConstants
    TextAlign2 = m_Align2
End Property
Public Property Let TextAlign2(ByVal Align As AlignmentConstants)
    If Align < vbLeftJustify Or Align > vbCenter Then
        ' should generate invalid property value error here
        Exit Property
    End If
    m_Align2 = Align
    PropertyChanged "Align2"
End Property

'/////// On-the-fly, dynamic, enumeration
Public Property Get Drives() As String
    Drives = m_Drive
End Property
Public Property Let Drives(ByVal theDrive As String)
    If InUserMode = True Then
        ' validate passed theDrive exists when in runtime
        ' if validation fails... Exit Property
    End If
    m_Drive = theDrive
    PropertyChanged "Drives"
End Property

'/////// On-the-fly, dynamic, enumeration - controls
Public Property Get ImageList() As Variant
    If InUserMode = True Then 'runtime, return object vs string
        Set ImageList = m_ImageList
    Else
        ImageList = m_sImageList
    End If
End Property
Public Property Let ImageList(ByVal theControl As Variant)
    ' you should validate the control passed is one that you are expecting
    If VarType(theControl) = vbString Then
        ' validate by string
        If theControl = "" Then
            Set m_ImageList = Nothing
            m_sImageList = vbNullString
        Else
            If ValidateTheControl(CStr(theControl), "ImageList", "", m_ImageList) Then
                m_sImageList = theControl
            Else
                ' should generate invalid property value error here
                Exit Property
            End If
        End If
    ' validate by control type
    ElseIf IsEmpty(theControl) Then
            Set m_ImageList = Nothing
            m_sImageList = vbNullString
    ElseIf VarType(theControl) = vbObject Then
        If TypeName(theControl) = "ImageList" Then
            If theControl.Index < 0 Then
                m_sImageList = theControl.Name
            Else
                m_sImageList = theControl.Name & "(" & theControl.Index & ")"
            End If
            Set m_ImageList = theControl
        Else
            ' should generate invalid property value error here
            Exit Property
        End If
    End If
    PropertyChanged "ImageList"
    PropertyChanged "ImageListCount" ' force sample property to update
End Property
Public Property Set ImageList(ByVal theControl As Variant)
    Me.ImageList = theControl ' allow the Let property to validate
End Property

'/////// Just sample prop to show that the ImageList property is working
Public Property Get ImageListCount() As Long
    If Not m_ImageList Is Nothing Then
        ImageListCount = m_ImageList.ListImages.Count
    End If
End Property
Public Property Let ImageListCount(ByVal newCount As Long)
    ' dummy property so this list shows in property sheet
End Property


'/////// Supporting routine for the usercontrol to set up property implementation
Private Sub ImplementProperties()
    ' initialize class & tell it which properties to implement
    If c_CustomProps Is Nothing Then
        
        Set c_CustomProps = New cCustomPropertyDisplay
        ' allow to run even though this control is uncompiled
        ' Not the best idea, but for this example, I want it run
        c_CustomProps.IgnoreIDESafety = True
        
        If c_CustomProps.Attach(Me, InUserMode) Then
            ' attach this property as a dropdown list of values
            c_CustomProps.AddProperty Me, "TextAlign2"
            c_CustomProps.AddProperty Me, "Drives"
            c_CustomProps.AddProperty Me, "ImageList"
            c_CustomProps.AddProperty Me, "FileName", False ' custom dialog
            c_CustomProps.AddProperty Me, "BackColor", False, True ' custom dialog
        Else
            Set c_CustomProps = Nothing ' failure; could be run-time vs design-time
            ' we don't want to run this in runtime
        End If
    End If
End Sub

'/////// IPropertyBrowserEvents callbacks
Private Sub IPropertyBrowserEvents_FormatPropertyDisplay(PropertyName As String, Display As String, Cancel As Boolean)
    ' specify exactly how you want the property sheet to display the value
    ' Any valid string is acceptable; however, VB's property sheet does not display unicode
    Select Case PropertyName
    Case "TextAlign2"                       ' should always display the same text as selected dropdown item
        If m_Align2 = vbLeftJustify Then
            Display = "0 - Left Aligned"
        ElseIf m_Align2 = vbRightJustify Then
            Display = "1 - Right Aligned"
        Else
            Display = "2 - Centered"
        End If
    Case "Drives"
        Display = m_Drives(m_IndexDrive)         ' set to selected dynamic enumeration's item
    Case "ImageList"
        Display = m_ImageLists(m_IndexImgLst) ' set to selected dynamic enumeration's item
    Case "BackColor"
        Display = FormatColor(UserControl.BackColor)
    Case Else
        Cancel = True                       ' not recommended; you should address each property implemented
    End Select
End Sub

Private Sub IPropertyBrowserEvents_FormatPropertyEnum(PropertyName As String, arrayDisplay() As String, arrayIndexes() As Variant)
    ' prepare the dropdown list text and their associated index value
    ' Any valid string is acceptable for arrayDisplay(); however, VB's property sheet does not display unicode
    Select Case PropertyName
    Case "TextAlign2"
        arrayDisplay() = Split("0 - Left Aligned,1 - Right Aligned,2 - Centered", ",") ' example of using Split to return arrayDisplay()
        arrayIndexes() = Array(vbLeftJustify, vbRightJustify, vbCenter)                ' example of using Array to return arrayIndexes()
    Case "Drives"
        Call GetDriveStrings            ' call local function to get current drive list
        arrayDisplay() = m_Drives()     ' send this back to the property sheet
        ' Option: we won't supply arrayIndexes and let the class supply them for us: 0 to UBound(m_Drives)
    Case "ImageList"
        Call GetParentControls("ImageList", m_ImageLists)
        arrayDisplay() = m_ImageLists()
        ' Option: we won't supply arrayIndexes and let the class supply them for us: 0 to UBound(m_Drives)
    End Select
End Sub

Private Sub IPropertyBrowserEvents_SetEnumPropertyValue(PropertyName As String, ItemData As Long, Value As Variant)
    ' user selected a value from the property sheet. Set the property's value here
    ' Value parameter is Variant and the value set must match the property type you declared
    Select Case PropertyName
    Case "TextAlign2"
        Value = ItemData
    Case "Drives"
        m_IndexDrive = ItemData                       ' update the index of the selected item
        If ItemData = 0& Then       ' if user selected "{None}" our property value is really ""
            Value = vbNullString
        Else
            Value = Left$(m_Drives(ItemData), 1) ' Else our value is really just the single character drive
        End If
    Case "ImageList"
        m_IndexImgLst = ItemData
        If ItemData = 0& Then       ' if user selected "{None}" our property value is Nothing
            Value = vbNullString
        Else
            Value = m_ImageLists(m_IndexImgLst)
        End If
    End Select
End Sub

Private Sub IPropertyBrowserEvents_ShowCustomPropertyPage(PropertyName As String)
    ' received only when property was implemented as a non-enumeration type property
    ' in this case, we want to display a custom dialog. Optionally, we can display
    ' a custom form (if it exists in the usercontrol) and show it modally.
    
    ' Any property changes must be made here, because we are bypassing VB's
    ' known methods of updating properties, VB doesn't know this is happening.
    
    On Error GoTo ExitRoutine
    Select Case PropertyName
    
    Case "FileName"
        ' setup your dialog or custom form as needed
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
            .FileName = vbNullString
            .DialogTitle = "Select Configuration File"
            .Filter = "Config Files|*.cfg|Text Files|*.txt|All Files|*.*"
            .FilterIndex = 0
        End With
        CommonDialog1.ShowOpen
        Me.FileName = CommonDialog1.FileName
        
    Case "BackColor"
        With CommonDialog1
            .CancelError = True
            .Flags = cdlCCRGBInit
            .Color = UserControl.BackColor
        End With
        CommonDialog1.ShowColor
        Me.BackColor = CommonDialog1.Color
    End Select
ExitRoutine:
    ' user canceled dialog
    If Err Then Err.Clear
End Sub

'/////// sample usercontrol events
Private Sub UserControl_Terminate()
    ' always include this in your usercontrol
    If Not c_CustomProps Is Nothing Then c_CustomProps.Detach
End Sub


Private Sub UserControl_Initialize()
    ReDim m_Drives(0)
    ReDim m_ImageLists(0)
    m_Drives(0) = "{None}"
    m_ImageLists(0) = "{None}"
End Sub

Private Sub UserControl_InitProperties()
    ' initialize any variables, then call this...
    ImplementProperties
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Align1 = PropBag.ReadProperty("Align1", vbLeftJustify)
    m_Align2 = PropBag.ReadProperty("Align2", vbLeftJustify)
    m_Drive = PropBag.ReadProperty("Drive", vbNullString)
    m_sImageList = PropBag.ReadProperty("ImgList", vbNullString)
    m_FileName = PropBag.ReadProperty("File", vbNullString)
    UserControl.BackColor = PropBag.ReadProperty("BkColor", vbButtonFace)
    ' With dynamic enums, we need to populate its list array immediately
    ' if the property will be implemented. This is because, VB calls
    ' the implementation's GetDisplayString to show it on the property sheet
    ' Additionally, we need to ensure the Index to that array is set properly
    If Len(m_Drive) Then
        GetDriveStrings ' verify the drive still exists on this system
        For m_IndexDrive = UBound(m_Drives) To 1 Step -1
            If Left$(m_Drives(m_IndexDrive), 1) = m_Drive Then Exit For
        Next
        If m_IndexDrive = 0& Then m_Drive = vbNullString ' no longer exists (think thumb drive)
    End If
    ' Another dynamic enumeration, but a little different
    If Len(m_sImageList) Then
        ' example of ensuring the control still exists on the parent
        If ValidateTheControl(m_sImageList, "ImageList", "IImageList", m_ImageList) = False Then
            m_sImageList = vbNullString
        Else
            ' we can't call GetParentControls because the controls may not yet exist. So to enable
            ' implementation's GetDisplayString to work, we simply fudge the array and its index
            ReDim Preserve m_ImageLists(0 To 1)
            m_ImageLists(1) = m_sImageList
            m_IndexImgLst = 1
            ' do whatever is needed when control is validated. Here we are updating ImageCount property
            PropertyChanged "ImageListCount"
        End If
    End If
    ' read properties then call this...
    ImplementProperties

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' write properties
    PropBag.WriteProperty "Align1", m_Align1, vbLeftJustify
    PropBag.WriteProperty "Align2", m_Align2, vbLeftJustify
    PropBag.WriteProperty "Drive", m_Drive, vbNullString
    PropBag.WriteProperty "ImgList", m_ImageLists(m_IndexImgLst), vbNullString
    PropBag.WriteProperty "BkColor", UserControl.BackColor, vbButtonFace
    PropBag.WriteProperty "File", m_FileName, vbNullString
End Sub



'/////// Supporting functions for the sample project only
Private Sub GetDriveStrings()
    ' get listing of logical drives
    Dim result As Long          ' Result of our api calls
    Dim strDrives As String     ' String to pass to api call
    
    ' Call GetLogicalDriveStrings with a buffer size of zero to
    ' find out how large our stringbuffer needs to be
    result = GetLogicalDriveStrings(0, strDrives)
    strDrives = String(result + 1, 0)
    ' Call again with our new buffer
    result = GetLogicalDriveStrings(result + 1, strDrives)
    m_Drives = Split("{None}" & vbNullChar & Left$(strDrives, InStr(strDrives, vbNullChar & vbNullChar) - 1), vbNullChar)
    
End Sub

Private Sub GetParentControls(ctrlType As String, ctrlNameArray() As String)
    ' get list of CtrlType controls on parent form
    On Error Resume Next
    Dim v As Object, sName As String
    Dim Count As Long, Index As Long, sortIndex As Long
    Index = UserControl.Parent.Controls.Count
    If Err Then
        Err.Clear
        Exit Sub
    End If
    For Each v In UserControl.Parent.Controls
        If TypeName(v) = ctrlType Then
            Count = Count + 1
            ReDim Preserve ctrlNameArray(0 To Count)
            Index = v.Index ' use error trapping to see if control is indexed
            If Index < 0 Then ' not indexed
                If Err Then Err.Clear
                sName = v.Name
            Else
                sName = v.Name & "(" & Index & ")"
            End If
            For Index = 1 To Count - 1
                ' might as well sort it too
                If sName < ctrlNameArray(Index) Then
                    For sortIndex = Count To Index + 1 Step -1
                        ctrlNameArray(sortIndex) = ctrlNameArray(sortIndex - 1)
                    Next
                    Exit For
                End If
            Next
            ctrlNameArray(Index) = sName
        End If
    Next

End Sub

Private Function ValidateTheControl(ctrlName As String, ctrlType As String, iFaceType As String, propVariable As Object) As Boolean
    ' customized validation of controls that are passed as string names
    ' If indexed, control name is formatted as:  ControlName(Index)
    
    ' ctrlName :: the formatted control name
    ' ctrlType :: is the control's type (i.e., ImageList, PictureBox, etc)
    '              To get this easily, in a click event: Debug.Print TypeName(Me.PictureBox1), TypeName(Me.ImageList1) etc
    ' iFaceType :: the interface of the control. Normally only needed during ReadProperties.
    '               Why? If the control isn't hosted on the form yet, TypeName() will return its interface vs its type.
    '               To get this easily, step thru the ValidateTheControl line in UserControl_ReadProperties.
    '               Then when you get to the Select Case statement below, open Debug window: ? TypeName(tObj)
    ' propVariable :: the variable that should be set to the control
    
    Dim tObj As Object, lPLoc As Long
    lPLoc = InStr(ctrlName, "(")
    On Error Resume Next
    If lPLoc Then
        Set tObj = UserControl.Parent.Controls(Left$(ctrlName, lPLoc - 1))(Val(Mid$(ctrlName, lPLoc + 1)))
    Else
        Set tObj = UserControl.Parent.Controls(ctrlName)
    End If
    If Err Then
        Err.Clear
    Else
        Select Case TypeName(tObj)
        Case ctrlType, iFaceType
            Set propVariable = tObj
            ValidateTheControl = True
        Case Else
        End Select
        Set tObj = Nothing
    End If

End Function

Private Function FormatColor(Color As Long) As String
    If Color < 0& Then
        FormatColor = "&H" & Hex(Color)
    Else
        FormatColor = "&H" & Right$("0000000" & Hex(Color), 8)
    End If
End Function

Private Function InUserMode() As Boolean
    ' support function to prevent "Client Site Not Available" errors
    ' in other applications that host usercontrols
    ' Example: Using MSAccess and compiled ocx
    '   Without this function...
    '       If any properties query Ambient.UserMode and control is added to an Access form: no problems
    '       If that control is then copied (i.e., CTRL+C): error
    '   With this function
    '       Error is trapped and Ambient.UserMode is assumed False since one can't CTRL+C controls in runtime
    
    On Error Resume Next
    InUserMode = Ambient.UserMode
    If Err Then Err.Clear
End Function


