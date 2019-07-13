VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "PrintDialog"
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
'Attribute VB_Name = "PrintDialog"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const PD_ALLPAGES = &H0 'Select the All Pages radio button.
Private Const PD_COLLATE = &H10 'Check the Collate check box. If this flag is set when the function returns, the user checked the box and the printer doesn't automatically support collation. If the box is checked and the printer does support it, this flag will not be set.
Private Const PD_DISABLEPRINTTOFILE = &H80000 'Disable the Print to File check box.
Private Const PD_ENABLEPRINTHOOK = &H1000 'Use the hook function pointed to by lpfnPrintHook to process the Print dialog box's messages.
Private Const PD_ENABLEPRINTTEMPLATE = &H4000 'Use the Print dialog box template specified by lpPrintTemplateName.
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000 'Use the preloaded Print dialog box template specified by hPrintTemplate.
Private Const PD_ENABLESETUPHOOK = &H2000 'Use the hook function pointed to by lpfnSetupHook to process the Print Setup dialog box's messages.
Private Const PD_ENABLESETUPTEMPLATE = &H8000 'Use the Print Setup dialog box template specified by lpSetupTemplateName.
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000 'Use the preloaded Print Setup dialog box template specified by hSetupTemplate.
Private Const PD_HIDEPRINTTOFILE = &H100000 'Hide the Print to File check box.
Private Const PD_NONETWORKBUTTON = &H200000 'Do not display any buttons associated with the network.
Private Const PD_NOPAGENUMS = &H8 'Disable the Page Range radio button and edit boxes.
Private Const PD_NOSELECTION = &H4 'Disable the Selection radio button.
Private Const PD_NOWARNING = &H80 'Do not warn the user if there is no default printer.
Private Const PD_PAGENUMS = &H2 'Select the Page Range radio button.
Private Const PD_PRINTSETUP = &H40 'Display the Print Setup dialog box instead of the Print dialog box.
Private Const PD_PRINTTOFILE = &H20 'Select the Print to File check box.
Private Const PD_RETURNDC = &H100 'Return a device context to the selected printer as hdc.
Private Const PD_RETURNDEFAULT = &H400 'Instead of displaying either dialog box, simply load information about the default printer into hDevMode and hDevNames. For this to work, those two values must be set to 0 before calling the function.
Private Const PD_RETURNIC = &H200 'Return an information context to the selected printer as hdc.
Private Const PD_SELECTION = &H1 ''Select the Selection radio button.
Private Const PD_SHOWHELP = &H800 'Display the Help button.
Private Const PD_USEDEVMODECOPIES = &H40000 'Same as Private Const PD_USEDEVMODECOPIESANDCOLLATE.
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000 'If the printer does not automatically support multiple copies or collation, disable the corresponding options in the dialog box. The number of copies to print and the collation setting will be placed into hDevMode. The information returned to this structure will specify the number of pages and the collation which the program must print with -- the printer will print the copies or collate itself.
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    HDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Declare Function apiPrintDlg Lib "comdlg32" Alias "PrintDlgA" (ByRef pPrintdlg As PRINTDLG_TYPE) As Long
'' Open a Print common dialog box.  Then display certain selections the user made,
'' such as the printer name, number of copies, and orientation.  Carefully note how memory
'' blocks are allocated to hold the two data structures containing information about the
'' printer device.  To save space, Visual Basic's Printer object, referring to the default
'' printer, is used to provide the defaults.  Of course, API functions could also be used
'' to get these defaults.
'Dim pd As PRINTDLG_TYPE ' holds information to make the dialog box
'Dim printmode As DEVMODE ' holds settings for the printer device
'Dim printnames As devnames ' holds device, driver, and port names
'Dim pMode As Long, pNames As Long  ' pointers to the memory blocks for the two DEV* structures
'Dim retval As Long  ' return value of function
'
'' First, load default settings into printmode.  Note that we only fill relevant information.
'printmode.dmDeviceName = Printer.DeviceName  ' name of the printer
'printmode.dmSize = Len(printmode)  ' size of the data structure
'printmode.dmFields = DM_ORIENTATION  ' identify which other members have information
'printmode.dmOrientation = DMORIENT_PORTRAIT  ' default to Portrait orientation
'
'' Next, load strings for default printer into printnames.  Note the unusual way in which such
'' information is stored.  This is explained on the DEVNAMES page.
'devnames.wDriverOffset = 8  ' offset of driver name string
'devnames.wDeviceOffset = devnames.wDriverOffset + 1 + Len(Printer.DriverName)  ' offset of printer name string
'devnames.wOutputOffset = devnames.wDeviceOffset + 1 + Len(Printer.port)  ' offset to output port string
'devnames.wDefault = 0  ' maybe this isn't the default selected printer
'' Load the three strings into the buffer, separated by null characters.
'devnames.extra = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.port & vbNullChar
'
'' Finally, load the initialization settings into pd, which is passed to the function.  We'll
'' set the pointers to the structures after this.
'pd.lStructSize = Len(pd)  ' size of structure
'pd.hwndOwner = Form1.hWnd  ' window Form1 is opening the Print dialog box
'' Flags: All Pages default, disable Print to File option, return device context:
'pd.flags = PD_ALLPAGES Or PD_DISABLEPRINTTOFILE Or PD_RETURNDC
'pd.nMinPage = 1  ' allow user to select first page of "document"
'pd.nMaxPage = 15  ' let's say there are 15 pages of the "document"
'' Note how we can ignore those members which will be set or are not used here.
'
'' Copy the data in printmode and printnames into the memory blocks we allocate.
'pd.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(printmode)  ' allocate memory block
'pMode = GlobalLock(pd.hDevMode)  ' get a pointer to the block
'CopyMemory ByVal pMode, printmode, Len(printmode)  ' copy structure to memory block
'retval = GlobalUnlock(pd.hDevMode)  ' unlock the block
'' Now do the same for printnames.
'pd.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(printnames)  ' allocate memory block
'pNames = GlobalLock(pd.hDevNames)  ' get a pointer to the block
'CopyMemory ByVal pNames, printnames, Len(printnames)  ' copy structure to memory block
'retval = GlobalUnlock(pd.hDevNames)  ' unlock the block
'
'' Finally, open the dialog box!
'retval = PrintDlg(pd)  ' looks so simple, doesn't it?
'
'' If the user hit OK, display some information about the selection.
'If retval <> 0 Then
'  ' First, we must copy the memory block data back into the structures.  This is almost identical
'  ' to the code above where we did the reverse.  Comments here are omitted for brevity.
'  pMode = GlobalLock(pd.hDevMode)
'  CopyMemory printmode, ByVal pMode, Len(printmode)
'  retval = GlobalUnlock(pd.hDevMode)
'  pNames = GlobalLock(pd.hDevNames)
'  CopyMemory printnames, ByVal pNames, Len(printnames)
'  retval = GlobalUnlock(pd.hDevNames)
'
'  ' Now, display the information we want.  We could instead use this info to print something.
'  Debug.Print "Printer Name: "; printmode.dmDeviceName
'  Debug.Print "Number of Copies:"; pd.nCopies
'  Debug.Print "Orientation: ";
'  If printmode.dmOrientation = DMORIENT_PORTRAIT Then Debug.Print "Portrait" Else Debug.Print "Landscape"
'  If (pd.flags And PD_SELECTION) = PD_SELECTION Then  ' user chose "Selection"
'    Debug.Print "Print the current selection."
'  ElseIf (pd.flags And PD_PAGENUMS) = PD_PAGENUMS Then  ' user chose a page range
'    Debug.Print "Print pages"; pd.nFromPage; "to"; pd.nToPage
'  Else  ' only one left is "All"
'    Debug.Print "Print all pages."
'  End If
'End If
'
'' No matter what, we have to deallocate the memory blocks from before.
'retval = GlobalFree(pd.hDevMode)
'retval = GlobalFree(pd.hDevNames)

