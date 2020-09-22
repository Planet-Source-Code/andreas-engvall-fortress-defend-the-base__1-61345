Attribute VB_Name = "Module1"
Public CannonHealth As Integer
Public CannonMax As Integer
Public CannonLevel As Integer
Public HouseHealth As Integer
Public HouseMax As Integer
Public HouseLevel As Integer

Public TimeLeft As Integer

Public Gold As Long
Public TotalScore As Long
Public Levels As Integer
Public Damage As Integer

Public Declare Function sndPlaySound& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) 'Declares the function so it can be used throughout the application
Public Const SND_SYNC& = &H0 'This will start the wave file, and doesn't allow any programs control until the wave is done. Don't use with SND_LOOP!
Public Const SND_ASYNC& = &H1 'This will start the wave file, and allow the calling program to continue.
Public Const SND_LOOP& = &H8 'This will continually loop the wave file endlessly (which is why you don't want to use SND_SYNC!)
Public Const SND_MEMORY& = &H4 'Play a wave file stored in memory.
Public Const SND_NODEFAULT& = &H2 'If the wave file can't be found for some reason, do not play any default wave files. It's a good idea to always use this.
Public Const SND_NOSTOP& = &H10 'Do not let this wave file interrupt another wave file being played.

