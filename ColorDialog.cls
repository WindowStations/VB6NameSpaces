VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "ColorDialog"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const CC_ANYCOLOR             As Long = &H100 'Allow the user to select any color.
Private Const CC_ENABLEHOOK           As Long = &H10  'Use the hook function specified by lpfnHook to process the Choose Color box's messages.
Private Const CC_ENABLETEMPLATE       As Long = &H20 'Use the dialog box template identified by hInstance and lpTemplateName.
Private Const CC_ENABLETEMPLATEHANDLE As Long = &H40  'Use the preloaded dialog box template identified by hInstance, ignoring lpTemplateName.
Private Const CC_FULLOPEN             As Long = &H2 'Automatically display the Define Custom Colors half of the dialog box.
Private Const CC_PREVENTFULLOPEN      As Long = &H4 'Disable the button that displays the Define Custom Colors half of the dialog box.
Private Const CC_RGBINIT              As Long = &H1 'Make the color specified by rgbResult be the initially selected color.
Private Const CC_SHOWHELP             As Long = &H8 'Display the Help button.
Private Const CC_SOLIDCOLOR           As Long = &H80 'Only allow the user to select solid colors. If the user attempts to select a non-solid color, convert it to the closest solid color.
Private Const GHND                    As Long = &H40 'Same as combining GMEM_MOVEABLE with GMEM_ZEROINIT.
Private Const GMEM_DDESHARE           As Long = &H2000 'Optimize the allocated memory for use in DDE conversations.
Private Const GMEM_DISCARDABLE        As Long = &H100 'Allocate discardable memory. (Cannot be combined with GMEM_FIXED.)
Private Const GMEM_FIXED              As Long = &H0 'Allocate fixed memory. The function's return value is a pointer to the beginning of the memory block. (Cannot be combined with GMEM_DISCARDABLE or GMEM_MOVEABLE.)
Private Const GMEM_MOVEABLE           As Long = &H2 'Allocate moveable memory. The memory block's lock count is initialized at 0 (unlocked). The function's return value is a handle to the beginning of the memory block. (Cannot be combined with GMEM_FIXED.)
Private Const GMEM_NOCOMPACT          As Long = &H10 'Do not compact any memory or discard any discardable memory to allocate the requested block.
Private Const GMEM_NODISCARD          As Long = &H20 'Do not discard any discardable memory to allocate the requested block.
Private Const GMEM_SHARE              As Long = &H2000 'Same as GMEM_DDESHARE.
Private Const GMEM_ZEROINIT           As Long = &H40 'Initialize the contents of the memory block to 0.
Private Const GPTR                    As Long = &H42 'Same as combining GMEM_FIXED with GMEM_ZEROINIT.
Private Type CHOOSECOLOR_TYPE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function apiChooseColor Lib "comdlg32" Alias "ChooseColorA" (ByRef lpcc As CHOOSECOLOR_TYPE) As Long
Private Declare Function apiGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function apiGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef source As Long, ByVal Length As Long)
Private Declare Function apiGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare Function apiGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private mvarAllowFullOpen  As Boolean
Private mvarAnyColor       As Boolean
Private mvarCustomColors() As Long
Private mvarFullOpen       As Boolean
Private mvarShowHelp       As Boolean
Private mvarSolidColorOnly As Boolean
Private mvarTag            As Object
Public mvarrgbResult       As Long
Friend Function ShowDialog() As DialogResult
    On Error Resume Next
    Dim cc                As CHOOSECOLOR_TYPE
    Dim hMem              As Long
    Dim pMem              As Long
    Dim C                 As Integer
    Dim ret               As Long
    Dim custcols(0 To 15) As Long
    hMem = apiGlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 64)
    pMem = apiGlobalLock(hMem)
    Call apiCopyMemory(ByVal pMem, custcols(0), 64)
    cc.lStructSize = Len(cc)
    With cc
        .hwndOwner = 0
        .hInstance = 0
        .rgbResult = 0 ' frmNameSpacesVB6.BackColor
        .lpCustColors = pMem
        .Flags = CC_ANYCOLOR Or CC_RGBINIT
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = ""
    End With
    ret = apiChooseColor(cc)
    If ret <> 0 Then
        Call apiCopyMemory(custcols(0), ByVal pMem, 64)
        mvarrgbResult = cc.rgbResult
        ShowDialog = IOK
    End If
    Call apiGlobalUnlock(hMem)
    Call apiGlobalFree(pMem)
End Function
Friend Sub Reset()
End Sub
Friend Sub Dispose()
End Sub

