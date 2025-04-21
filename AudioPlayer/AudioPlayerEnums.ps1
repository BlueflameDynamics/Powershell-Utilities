$MyEnums = @(
[CheckboxID],
[FileMenuItem],
[HelpMenuItem],
[ImageListID],
[LvwColumn],
[LvwCtxItem],
[MainMenuItem],
[MediaButton],
[MediaIcon],
[ToolMenuItem],
[RegistryKey]
)
#region Custom Enums used in Audioplayer.ps1
Enum CheckboxID{
	Loop
	AutoClose
	Recurse
}

Enum FileMenuItem{
	OpenPlaylist
	NewPlaylist
	EditPlaylist
	ReloadPlaylist
	Find
	FindNext
	Exit
}

Enum HelpMenuItem{
	Help
	About
	HostInformation
}

Enum ImageListID{
	SmIcon
	LgIcon
}

Enum LvwColumn{
	Icon
	Duration
	File
}

Enum LvwCtxItem{
	OpenPlaylist
	NewPlaylist
	EditPlaylist
	ReloadPlaylist
	Find
	FindNext
	FontSettings
	Properties
	Exit
}

Enum MainMenuItem{
	File
	Tools
	Help
}

Enum MediaButton{
	Play
	Pause
	Stop
}

Enum MediaIcon{
	AudioFile
	Directory
}

Enum ToolMenuItem{
	ResetColumnWidth
	FontSettings
	LockVolume
	SaveSettings
	DeleteSettings
}

Enum RegistryKey{
	Default = 0
	AutoClose
	AutoPlay
	HelpRtbFont
	HideVolumeLock
	LockVolume
	LoopPlayback
	MainFormSize
	MainLvwColumnWidth
	MainLvwFont
	Minimized
	MiniMode
	Playlist
	RecurseDirectory
	Volume
}
#endregion