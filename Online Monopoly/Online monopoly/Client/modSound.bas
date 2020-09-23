Attribute VB_Name = "modSound"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal parent As Long, ByVal operation As String, ByVal file As String, ByVal parameters As String, ByVal directory As String, ByVal mode As Long) As Long

Public v_dx As New DirectX7
Public v_dmp As DirectMusicPerformance
Public v_dml As DirectMusicLoader
Public v_dms As DirectMusicSegment
Public v_dmss As DirectMusicSegmentState
Public vs_filename As String
Public vl_volume As Long

Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ALIAS = &H10000 'lpszName is a string identifying the name of the system event sound to play.
Public Const SND_ALIAS_ID = &H110000 'lpszName is a string identifying the name of the predefined sound identifier to play.
Public Const SND_APPLICATION = &H80 'lpszName is a string identifying the application-specific event association sound to play.
Public Const SND_ASYNC = &H1 'Play the sound asynchronously -- return immediately after beginning to play the sound and have it play in the background.
Public Const SND_FILENAME = &H20000 'lpszName is a string identifying the filename of the .wav file to play.
Public Const SND_LOOP = &H8 'Continue looping the sound until this function is called again ordering the looped playback to stop. SND_ASYNC must also be specified.
Public Const SND_MEMORY = &H4 'lpszName is a numeric pointer refering to the memory address of the image of the waveform sound loaded into RAM.
Public Const SND_NODEFAULT = &H2 'If the specified sound cannot be found, terminate the function with failure instead of playing the SystemDefault sound. If this flag is not specified, the SystemDefault sound will play if the specified sound cannot be located and the function will return with success.
Public Const SND_NOSTOP = &H10 'If a sound is already playing, do not prematurely stop that sound from playing and instead return with failure. If this flag is not specified, the playing sound will be terminated and the sound specified by the function will play instead.
Public Const SND_NOWAIT = &H2000 'If a sound is already playing, do not wait for the currently playing sound to stop and instead return with failure.
Public Const SND_PURGE = &H40 'Stop playback of any waveform sound. lpszName must be an empty string.
Public Const SND_RESOURCE = &H4004 'lpszName is the numeric resource identifier of the sound stored in an application. hModule must be specified as that application's module handle.
Public Const SND_SYNC = &H0 'Play the sound synchronously -- do not return until the sound has finished playing.

Public playSoundBool As Boolean
Public playMusicBool As Boolean
Public musicVolume As Long
Dim a As Variant

Public Sub PlayMidi(NumMid As Integer)
    Set v_dml = v_dx.DirectMusicLoaderCreate
    Set v_dmp = v_dx.DirectMusicPerformanceCreate
    
    Call v_dmp.Init(Nothing, frmMain.hWnd)
    Call v_dmp.SetPort(-1, 1)
    
    If NumMid = 0 Then
        Randomize
        NumMid = (Rnd() * 100 Mod 5) + 2
    End If
    vs_filename = "music" & NumMid & ".mid"
    v_dml.SetSearchDirectory (App.Path & "\midi")
    Set v_dms = v_dml.LoadSegment(vs_filename)
    If StrConv(Right(vs_filename, 4), vbLowerCase) = ".mid" Then
        v_dms.SetStandardMidiFile
    End If
    
    Call v_dmp.SetMasterAutoDownload(True)
    Call v_dms.Download(v_dmp)
    Call v_dmp.SetMasterVolume(musicVolume)
    Set v_dmss = v_dmp.PlaySegment(v_dms, 0, 0)
    frmMain.tmrMusic.Enabled = True
End Sub

Public Function play(filename As String)
    On Error Resume Next
    If playSoundBool Then
        a = PlaySound(App.Path & "/sounds/" & filename, 0, SND_FILENAME Or SND_ASYNC)
    End If
End Function

Public Sub CloseMidi()
    If v_dms Is Nothing Then Exit Sub
    Call v_dmp.Stop(v_dms, v_dmss, 0, 0)
    Call v_dms.Unload(v_dmp)
    frmMain.tmrMusic.Enabled = False
End Sub
