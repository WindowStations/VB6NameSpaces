VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Audio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SND_APPLICATION As Long = &H80 '  look for application specific association
Private Const SND_ALIAS       As Long = &H10000 '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID    As Long = &H110000 '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC       As Long = &H1 '  play asynchronously
Private Const SND_FILENAME    As Long = &H20000 '  name is a file name
Private Const SND_LOOP        As Long = &H8 '  loop the sound until next sndPlaySound
Private Const SND_MEMORY      As Long = &H4 '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT   As Long = &H2 '  silence not default, if sound not found
Private Const SND_NOSTOP      As Long = &H10 '  don't stop any currently playing sound
Private Const SND_NOWAIT      As Long = &H2000 '  don't wait if the driver is busy
Private Const SND_PURGE       As Long = &H40 '  purge non-static events for task
Private Const SND_RESOURCE    As Long = &H40004 '  name is a resource name or atom
Private Const SND_SYNC        As Long = &H0 '  play synchronously (default)
Private Declare Function apiPlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Friend Function Play(ByVal fPath As String) As Long
    Play = apiPlaySound(fPath, ByVal 0, SND_FILENAME Or SND_ASYNC)
End Function
Friend Function Stop_() As Long
    Stop_ = apiPlaySound(vbNullString, 0, SND_PURGE)
End Function
