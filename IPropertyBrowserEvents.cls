VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPropertyBrowserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' This interface must be used vs a WithEvents cCustomPropertyDisplay solution.
' VB can block/halt RaiseEvent calls when a modal window is displayed in IDE
' and since some properties can display modal windows in the form of a property page,
' picture selection dialog, and/or color dialog, the property sheet can
' temporarily lose its display values. Only an issue when class is used with
' uncompiled usercontrols.

Option Explicit

' Add the following in your usercontrol's declarations section:  Implements IPropertyBrowserEvents
' Each of these must have some code inside of your usercontrol. If no code is applicable, add a tick mark

' More notes:  Which must be responded to and which do not?
' When you called cCustomPropertyPages:AddProperty, depending on the Boolean parameters...
'   If Enumeration=True, respond to the following
'       FormatPropertyDisplay, FormatPropertyEnum, SetEnumPropertyValue
'       Note: LockIDEDisplay is ignored
'   If Enumeration=False And LockIDEDisplay=False, respond to the following
'       ShowCustomPropertyPage
'   If Enumeration=False And LockIDEDisplay=True, respond to the following
'       ShowCustomPropertyPage, FormatPropertyDisplay


Public Sub FormatPropertyEnum(PropertyName As String, arrayDisplay() As String, arrayIndexes() As Variant)
    ' this event provides the property page the dropdown, formatted, listing
    
    'Parameters...
    '   PropertyName: the name you added via cCustomPropertyDisplay:AddProperty call
    '   arrayDisplay(): an array of strings that appear in the combobox presented to the user
    '   arrayIndexes(): an array of numeric, non-decimal numbers only. Similar to combobox's ItemData property
    '           The indexes can be the actual Long values of the combobox items, or
    '           The indexes can just be indexes of 0 thru UBound(arrayDisplay)-1
    '       These are used in the SetEnumPropertyValue event as you will see below
    ' Note: If not provided, arrayIndexes() will default with values of LBound(arrayDisplay) to UBound(arrayDiplay)
    ' Note: If provided, both arrayDisplay & arrayIndexes must be zero-LBound. Both must have same UBound (my requirement)
    ' Note: arrayIndexes() is Variant vs Long to aid the coder in creating quick arrays via VB's Array() function
    ' Note: arrayDisplay() can contain spaces or any characters; VB property sheet & dropdown is not unicode compatible.
    ' Note: arrayDisplay() is not required to have a numerical prefix in front of each list item; however, it should
    '       be used when the listing is an enumeration so user knows the list item's constant value. It should not be
    '       used for on-the-fly, dynamic, enumerations, since your property is not referencing an actual enumeration
    
    ' Here are some examples of usage:
    ' -----------------------------------------
'        On Error GoTo ExitRoutine
'        Select Case PropertyName
'        Case "Alignment" ' we don't want 0 - vbLeftJustify, 1 - vbRightJustify, 2 - vbCenter showing in combobox
'                         ' we want       0 - Left Aligned, 1 - Right Aligned, 2 - Centered
'            arrayDisplay() = Split("0 - Left Aligned,1-Right Aligned,2 - Centered", ",")
'            arrayIndexes = Array(vbLeftJustify, vbRightJustify, vbCenter)
'
'        ' following example enables a string-like, dynamic, enumeration
'        Case "Drive"
'            Call GetDriveList(m_drives()) ' call your local function to load the m_Drives() array with the formatted display values
'            arrayDisplay = m_drives()
'            ' Option: don't provide arrayIndexes and it will be filled as:  LBound(m_Drives) To UBound(m_Drives)
'        End Select
'
'        ExitRoutine:
'        If Err Then
'            ' do something about the error?
'        End If
End Sub

Public Sub SetEnumPropertyValue(PropertyName As String, ItemData As Long, Value As Variant)
    ' this event sets the property value when a user selects an item from the dropdown list or
    ' double clicks on the property name.
    
    ' Parameters...
    '   PropertyName: the name you added via cCustomPropertyDisplay:AddProperty call
    '   ItemData: the arrayIndexes() value of the selected combobox item. See FormatPropertyEnum above
    '   Value: the actual value that the property should be set to
    '       Note: Ensure you use Value correctly for your property
    '       If property value is an object, use: Set Value = [whatever]
    '       If property value is a Byte, Integer, Single, Double, Date, Currency, etc, use the appropriate Conversion function: CDbl, CDate, etc
    '       If property value is a string, simply: Value = [string]
    
    ' Here are some examples of usage:
    ' -----------------------------------------
'        On Error GoTo ExitRoutine
'        Select Case PropertyName
'        Case "Alignment"        ' when actual Long values are supplied to FormatPropertyEnum, you can use the ItemData
'            Value = ItemData
'        Case "Drive"
'            m_DriveIndex = ItemData ' m_DriveIndex would be a reference you keep to know which m_Drives() array item is selected
'            Value = m_Drives(ItemData)
'        End Select
'
'        ExitRoutine:
'        If Err Then
'            ' do something about the error
'        End If
End Sub


Public Sub FormatPropertyDisplay(PropertyName As String, Display As String, Cancel As Boolean)
    ' this event sets the text to be displayed in the property sheet for the property
    ' The display text must match the selected dropdown value exactly for best results, otherwise,
    '   if user double clicks property name, VB will not select the next item in the listing.
    
    ' Parameters...
    '   PropertyName: the name you added via cCustomPropertyDisplay:AddProperty call
    '   Display: the string you want displayed in the property sheet
    '   Cancel: set to true if you don't want to display anything custom, then VB will display it's value instead
    '       Not recommended. This can break the property name double click behavior mentioned above
    
    ' Here are some examples of possible usage:
    ' -----------------------------------------
'        On Error GoTo ExitRoutine
'        Select Case PropertyName
'        Case "Alignment"        ' assuming you have m_Align variable containing the following possible values
'            If m_Align = vbLeftJustify Then
'                Display = "0 - Left Aligned"
'            ElseIf m_Align = vbRightJustify Then
'                Display = "1 - Right Aligned"
'            Else
'                Display = "2 - Centered"
'            End If
'        Case "Drive"            ' assuming you have m_Drives() string with entries like "0 - A:\", "1 - C:\", "2 - D:\", etc
'            Display m_Drives(m_DriveIndex)
'        Case Else
'            Cancel = True       ' not recommended; see above
'        End Select
'
'        ExitRoutine:
'        If Err Then
'            ' do something about the error?
'            Cancel = True
'        End If
End Sub

Public Sub ShowCustomPropertyPage(PropertyName As String)
    ' this event prompts you to display whatever custom dialog/form you wish to use instead of a property page
    ' Within this event you must set the property's value as necessary. VB will not know the property changed

    ' Parameters...
    '   PropertyName: the name you added via cCustomPropertyDisplay:AddProperty call
    
    ' Here are some examples of possible usage:
    ' -----------------------------------------
'        On Error GoTo ExitRoutine
'        Select Case PropertyName
'        Case "FileName"
'               CommonDialog1.CancelError = True
'               CommonDialog1.ShowOpen
'               Me.FileName = CommonDialog1.FileName
'        Case Else
'
'        End Select
'
'        ExitRoutine:
'        If Err Then
'            ' do something about the error?
'        End If

End Sub

'///////////////////////////////////////////////////////////////////
' function provided for your convenience & is not part of this class
' Just added here so you always have a copy should you need it.
'///////////////////////////////////////////////////////////////////
'Private Function InUserMode() As Boolean
'    ' support function to prevent "Client Site Not Available" errors
'    ' in other applications that host usercontrols
'    ' Example: Using MSAccess and compiled ocx
'    '   Without this function...
'    '       If any properties query Ambient.UserMode and control is added to an Access form: no problems
'    '       If that control is then copied (i.e., CTRL+C): error
'    '   With this function
'    '       Error is trapped and Ambient.UserMode is assumed False since one can't CTRL+C controls in runtime
'
'    On Error Resume Next
'    InUserMode = Ambient.UserMode
'    If Err Then Err.Clear
'End Function
'///////////////////////////////////////////////////////////////////
