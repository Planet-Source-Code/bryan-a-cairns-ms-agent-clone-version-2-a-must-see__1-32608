Attribute VB_Name = "mod_Wavs"
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
     (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'  flag values for uFlags parameter
Const SND_SYNC = &H0            '  play synchronously (default)
Const SND_ASYNC = &H1           '  play asynchronously
Const SND_NODEFAULT = &H2       '  silence not default, if sound not found
Const SND_MEMORY = &H4          '  lpszSoundName points to a memory file
Const SND_ALIAS = &H10000       '  name is a WIN.INI [sounds] entry
Const SND_FILENAME = &H20000    '  name is a file name
Const SND_RESOURCE = &H40004    '  name is a resource name or atom
Const SND_ALIAS_ID = &H110000   '  name is a WIN.INI [sounds] entry identifier
Const SND_ALIAS_START = 0       '  must be > 4096 to keep strings in same section of resource file
Const SND_LOOP = &H8            '  loop the sound until next sndPlaySound
Const SND_NOSTOP = &H10         '  don't stop any currently playing sound
Const SND_VALID = &H1F          '  valid flags          / ;Internal /
Const SND_NOWAIT = &H2000       '  don't wait if the driver is busy
Const SND_VALIDFLAGS = &H17201F '  Set of valid flag bits.  Anything outside
                                '  this range will raise an error
Const SND_RESERVED = &HFF000000 '  In particular these flags are reserved
Const SND_TYPE_MASK = &H170007


Public Sub LoopAWav(sFile As String)
On Error GoTo EH
'Loops a Sound
    Dim L As Long
    L = sndPlaySound(sFile, SND_LOOP Or SND_ASYNC)
    If L = 0 Then
        If Dir$(sFile) = "" Then
            MsgBox "Unable to play the sound! File not found!"
        Else
            'MsgBox "Unable to play the sound!"
        End If
    End If
Exit Sub
EH:
MsgBox Err.Number, vbCritical, "Loop Wav"
Exit Sub
End Sub

Public Sub PlayAWav(sFile As String)
On Error GoTo EH
'Plays a WAV File
    Dim L As Long
    L = sndPlaySound(sFile, SND_ASYNC Or SND_NODEFAULT)
    If L = 0 Then
        If Dir$(sFile) = "" Then
            MsgBox "Unable to play the sound! File not found!"
        Else
            'MsgBox "Unable to play the sound!"
        End If
    End If
Exit Sub
EH:
MsgBox Err.Number, vbCritical, "Play Wav"
Exit Sub
End Sub

Public Sub StopAllWavs()
'Stops a ALL Wav Files that are playing
    Dim L As Long
    L = sndPlaySound("", SND_SYNC Or SND_NODEFAULT)
End Sub







