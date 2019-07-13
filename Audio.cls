VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Audio"
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
'Attribute VB_Name = "Audio"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
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
    Play = apiPlaySound(fPath, 0, SND_FILENAME Or SND_ASYNC)
End Function
Friend Function Stop_() As Long
    Stop_ = apiPlaySound(vbNullString, 0, SND_PURGE)
End Function
Friend Sub PlaySync(ByRef paths() As String) 'Play songs syncronously one after the other is finished
    Dim i As Long
    For i = 0 To UBound(paths)
        WaitToComplete (paths(i))
    Next
    '  WaitToComplete ("C:\SoundFile2.wav")
    '  WaitToComplete ("C:\SoundFile3.wav")
End Sub
Private Sub WaitToComplete(ByVal sPath As String) 'Wait up to an hour
    Dim i As Long
    For i = 1 To 900
        If apiPlaySound(sPath, 0, SND_FILENAME Or SND_ASYNC Or SND_NOSTOP) <> 0 Then 'If started
            If i > 1 Then  'If finished and now playing a second time
                Call apiPlaySound(vbNullString, 0, SND_FILENAME Or SND_PURGE) 'Stop song from playing twice
                Exit For
            End If
        End If
        Threading.Thread.Sleep (3000) 'sleep 4 seconds
        Application.DoEventsQueued
    Next
End Sub

