VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "SaveFileDialog"
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
'Attribute VB_Name = "SaveFileDialog"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Const OFN_ALLOWMULTISELECT     As Long = &H200
Private Const OFN_CREATEPROMPT         As Long = &H2000
Private Const OFN_ENABLEHOOK           As Long = &H20
Private Const OFN_ENABLETEMPLATE       As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER             As Long = &H80000 '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT   As Long = &H400
Private Const OFN_FILEMUSTEXIST        As Long = &H1000
Private Const OFN_HIDEREADONLY         As Long = &H4
Private Const OFN_LONGNAMES            As Long = &H200000
Private Const OFN_NOCHANGEDIR          As Long = &H8
Private Const OFN_NODEREFERENCELINKS   As Long = &H100000
Private Const OFN_NOLONGNAMES          As Long = &H40000
Private Const OFN_NONETWORKBUTTON      As Long = &H20000
Private Const OFN_NOREADONLYRETURN     As Long = &H8000& '*see comments
Private Const OFN_NOTESTFILECREATE     As Long = &H10000
Private Const OFN_NOVALIDATE           As Long = &H100
Private Const OFN_OVERWRITEPROMPT      As Long = &H2
Private Const OFN_PATHMUSTEXIST        As Long = &H800
Private Const OFN_READONLY             As Long = &H1
Private Const OFN_SHAREAWARE           As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH     As Long = 2
Private Const OFN_SHAREWARN            As Long = 0
Private Const OFN_SHARENOWARN          As Long = 1
Private Const OFN_SHOWHELP             As Long = &H10
Private Const OFN_ENABLESIZING         As Long = &H800000
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function apiGetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (ByRef lpofn As OPENFILENAME) As Long
Private mvarAddExtension                 As Boolean
Private mvarAutoUpgradeEnabled           As Boolean
Private mvarCheckFileExists              As Boolean
Private mvarCheckPathExists              As Boolean
Private mvarDefaultExt                   As String
Private mvarDereferenceLinks             As Boolean
Private mvarFileName                     As String
Private mvarFileNames()                  As String
Private mvarFilter                       As String
Private mvarFilterIndex                  As Long
Private mvarInitialDirectory             As String
Private mvarReadOnlyChecked              As Boolean
Private mvarRestoreDirectory             As Boolean
Private mvarSafeFileName                 As String
Private mvarSafeFileNames()              As String
Private mvarShowHelp                     As Boolean
Private mvarShowReadOnly                 As Boolean
Private mvarSupportMultiDottedExtensions As Boolean
Private mvarTag                          As Object
Private mvarTitle                        As String
Private mvarValidateNames                As Boolean
Friend Property Let AddExtension(ByVal vData As Boolean)
    mvarAddExtension = vData
End Property
Friend Property Get AddExtension() As Boolean
    AddExtension = mvarAddExtension
End Property
Friend Property Let AutoUpgradeEnabled(ByVal vData As Boolean)
    mvarAutoUpgradeEnabled = vData
End Property
Friend Property Get AutoUpgradeEnabled() As Boolean
    AutoUpgradeEnabled = mvarAutoUpgradeEnabled
End Property
Friend Property Let CheckFileExists(ByVal vData As Boolean)
    mvarCheckFileExists = vData
End Property
Friend Property Get CheckFileExists() As Boolean
    CheckFileExists = mvarCheckFileExists
End Property
Friend Property Let CheckPathExists(ByVal vData As Boolean)
    mvarCheckPathExists = vData
End Property
Friend Property Get CheckPathExists() As Boolean
    CheckPathExists = mvarCheckPathExists
End Property
Friend Property Let DefaultExt(ByVal vData As String)
    mvarDefaultExt = vData
End Property
Friend Property Get DefaultExt() As String
    DefaultExt = mvarDefaultExt
End Property
Friend Property Let DereferenceLinks(ByVal vData As Boolean)
    mvarDereferenceLinks = vData
End Property
Friend Property Get DereferenceLinks() As Boolean
    DereferenceLinks = mvarDereferenceLinks
End Property
Friend Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Friend Property Get FileName() As String
    FileName = mvarFileName
End Property
Friend Property Let FileNames(ByVal vData As Variant)
    mvarFileNames = vData
End Property
Friend Property Get FileNames() As Variant
    FileNames = mvarFileNames
End Property
Friend Property Let Filter(ByVal vData As String)
    mvarFilter = vData
End Property
Friend Property Get Filter() As String
    Filter = mvarFilter
End Property
Friend Property Let FilterIndex(ByVal vData As Long)
    mvarFilterIndex = vData
End Property
Friend Property Get FilterIndex() As Long
    FilterIndex = mvarFilterIndex
End Property
Friend Property Let InitialDirectory(ByVal vData As String)
    mvarInitialDirectory = vData
End Property
Friend Property Get InitialDirectory() As String
    InitialDirectory = mvarInitialDirectory
End Property
Friend Property Let ReadOnlyChecked(ByVal vData As Boolean)
    mvarReadOnlyChecked = vData
End Property
Friend Property Get ReadOnlyChecked() As Boolean
    ReadOnlyChecked = mvarReadOnlyChecked
End Property
Friend Property Let RestoreDirectory(ByVal vData As Boolean)
    mvarRestoreDirectory = vData
End Property
Friend Property Get RestoreDirectory() As Boolean
    RestoreDirectory = mvarRestoreDirectory
End Property
Friend Property Let SafeFileName(ByVal vData As String)
    mvarSafeFileName = vData
End Property
Friend Property Get SafeFileName() As String
    SafeFileName = mvarSafeFileName
End Property
Friend Property Let SafeFileNames(ByVal vData As Variant)
    mvarSafeFileNames = vData
End Property
Friend Property Get SafeFileNames() As Variant
    SafeFileNames = mvarSafeFileNames
End Property
Friend Property Let ShowHelp(ByVal vData As Boolean)
    mvarShowHelp = vData
End Property
Friend Property Get ShowHelp() As Boolean
    ShowHelp = mvarShowHelp
End Property
Friend Property Let ShowReadOnly(ByVal vData As Boolean)
    mvarShowReadOnly = vData
End Property
Friend Property Get ShowReadOnly() As Boolean
    ShowReadOnly = mvarShowReadOnly
End Property
Friend Property Let SupportMultiDottedExtensions(ByVal vData As Boolean)
    mvarSupportMultiDottedExtensions = vData
End Property
Friend Property Get SupportMultiDottedExtensions() As Boolean
    SupportMultiDottedExtensions = mvarSupportMultiDottedExtensions
End Property
Friend Property Let Tag(ByVal vData As Object)
    mvarTag = vData
End Property
Friend Property Get Tag() As Object
    Tag = mvarTag
End Property
Friend Property Let title(ByVal vData As String)
    mvarTitle = vData
End Property
Friend Property Get title() As String
    title = mvarTitle
End Property
Friend Property Let ValidateNames(ByVal vData As Boolean)
    mvarValidateNames = vData
End Property
Friend Property Get ValidateNames() As Boolean
    ValidateNames = mvarValidateNames
End Property
Friend Function ShowDialog() As DialogResult
    Dim ofn   As OPENFILENAME
    Dim fname As String
    Dim ret   As Long
    With ofn
        .lStructSize = Len(ofn)
        If Me.FileName <> "" Then .lpstrFile = Me.FileName & Space(260 - Len(Me.FileName)) & vbNullChar
        If Me.FileName = "" Then .lpstrFile = Space(260) & vbNullChar
        .lpstrInitialDir = Me.InitialDirectory '= "C:\" & vbNullChar
        .lpstrTitle = Me.title & vbNullChar ' "Select a File" & vbNullChar
        .lpstrFilter = Me.Filter '"Text Files (*.txt)" & vbNullChar & "*.txt" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nFilterIndex = Me.FilterIndex
        If Me.CheckFileExists = True Then .Flags = .Flags Or OFN_FILEMUSTEXIST
        If Me.CheckPathExists = True Then .Flags = .Flags Or OFN_PATHMUSTEXIST
        If Me.ShowHelp = True Then .Flags = .Flags Or OFN_SHOWHELP
        If Me.ReadOnlyChecked = True Then .Flags = .Flags Or OFN_READONLY
        If Me.ReadOnlyChecked = False Then .Flags = .Flags Or OFN_HIDEREADONLY
        If Me.DereferenceLinks = False Then .Flags = .Flags Or OFN_NODEREFERENCELINKS
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(260) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .nMaxCustomFilter = 0
        .hwndOwner = 0
        .hInstance = 0
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
        '        .lStructSize = Len(ofn)
        '        .hwndOwner = 0
        '        .hInstance = 0
        '        .lpstrFilter = "Text Files (*.txt)" & vbNullChar & "*.txt" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        '        .nMaxCustomFilter = 0
        '        .nFilterIndex = 1
        '        .lpstrFile = Space(260) & vbNullChar
        '        .nMaxFile = Len(.lpstrFile)
        '        .lpstrFileTitle = Space(260) & vbNullChar
        '        .nMaxFileTitle = Len(.lpstrFileTitle)
        '        .lpstrInitialDir = "C:\" & vbNullChar
        '        .lpstrTitle = "Select a File" & vbNullChar
        '        .Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        '        .nFileOffset = 0
        '        .nFileExtension = 0
        '        .lCustData = 0
        '        .lpfnHook = 0
    End With
    ret = apiGetSaveFileName(ofn)
    ShowDialog = ret
    If ret = DialogResult.IOK Then
        mvarFileName = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        mvarFileName = ""
    End If
End Function
Friend Function Reset()
    Me = New OpenFileDialog
    With Me
        .AddExtension = True
        .AutoUpgradeEnabled = True
        .CheckFileExists = True
        .CheckPathExists = True
        .DefaultExt = ".txt"
        .FileName = ""
        .InitialDirectory = ""
        .title = ""
        .Tag = Nothing
        .Filter = ""
        .FilterIndex = 1
        'todo
    End With
End Function
Friend Function Dispose()
    Me = Nothing
End Function

