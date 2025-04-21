<#
.NOTES
-------------------------------------
Name:	AudioPlayer.ps1
Version: 1.0u - 04/04/2025
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This script launches Audio files from a definable location by use of a WinForm & Playlist.
------------------------------------------------------------------------------------------
Supported Audio File Types are: .acc, .aif, .aiff, .au, .m4a, .mp3,.snd, .wav, .wma
------------------------------------------------------------------------------------------

.DESCRIPTION

To Run the script without a console window use the included
PsRun3.exe/PsRun3X64.exe PowerShell Launcher

Run PsRun3.exe/PsRun3X64.exe with a -help parameter to display syntax and sample usage.

%PowerShellDevLib%\PsRun3.exe "%PowerShellDevLib%\AudioPlayer.ps1"
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
.Parameter PlayList - Alias: Play
A text file with one entry per line naming a file or a directory of files to play.

.Parameter ErrorLog - Alias: Elog
Use to set the Error Log file, default: .\PS_Audio_Player_Errors.txt

.Parameter Volume - Alias: Vol
Used to set player volume, Values 0-100, default: -1 (use current)

.PARAMETER FontName - Alias: Fn
Name of the font to be used.

.PARAMETER FontSize - Alias: Fz
Size of the font to be used, between 9-24 points.

.PARAMETER FontStyle - Alias: Fs
Style of the font to be used: Bold, Italic, BoldItalic, Regular.

.PARAMETER Recurse - Alias: R
Use to recurse directories within the playlist.

.PARAMETER LoopPlayback - Alias: Loop
Use to Loop Playback of Playlist

.PARAMETER AutoPlay - Alias: Auto+
Use to Automatticly Play Command-Line Playlist on Load.

.PARAMETER AutoClose - Alias: Auto-
Use to Automattically Close the Player after completing Playlist.
Overrides -LoopPlayback

.PARAMETER LockVolume - Alias: LockV
Locks system audio volume

.PARAMETER HideLockVolume - Alias: LockH
Hides system audio volume lock

.PARAMETER LockSettings - Alias: LS
Used to Lock/Unlock registry settings, default: '' (Not Set)

.PARAMETER SaveSettings - Alias: Save
Used to Save registry settings

.PARAMETER MiniMode - Alias: Mini
Causes the main ListView to be hidden and resizes the form for minimum GUI.

.PARAMETER Minimized - Alias: M
Causes the main window to be Minimized to the Taskbar.

.EXAMPLE
PS> .\AudioPlayer.ps1 -Play "\\MediaServer01\Public\Shared Music\Playlists\NewOrleansJazz.apl" 
-Fn "Times New Roman" -Fz 16 -Fs Bold -Mini
#>

[CmdletBinding()]
param(
	[Parameter()][Alias('Play')][String]$PlayList = '',
	[Parameter()][Alias('ELog')][String]$ErrorLog = '.\PS_Audio_Player_Errors.txt',
	[Parameter()][Alias('Vol')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(0,100)]
		[Int]$Volume = -1, #Indicates value not set
	[Parameter()][Alias('Fn')][String]$FontName = 'Lucida Console',
	[Parameter()][Alias('Fz')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(9,24)]
		[Int]$FontSize = 12,
	[Parameter()][Alias('Fs')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Bold','Italic','BoldItalic','Regular')]
		[String]$FontStyle = 'Regular',
	[Parameter()][Alias('LS')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Lock','Unlock')]
		[String]$LockSettings = '', #Indicates value not set
	[Parameter()][Alias('R')][Switch]$Recurse,
	[Parameter()][Alias('Loop')][Switch]$LoopPlayback,
	[Parameter()][Alias('LockV')][Switch]$LockVolume,
	[Parameter()][Alias('LockH')][Switch]$HideLockVolume,
	[Parameter()][Alias('Auto+')][Switch]$AutoPlay,
	[Parameter()][Alias('Auto-')][Switch]$AutoClose,
	[Parameter()][Alias('Save')][Switch]$SaveSettings,
	[Parameter()][Alias('Mini')][Switch]$MiniMode,
	[Parameter()][Alias('M')][Switch]$Minimized)

#Requires -Version 5

#region Module Import
Import-Module -Name .\AppRegistry.ps1 -Force
Import-Module -Name .\AudioPlayerEnums.ps1 -Force
Import-Module -Name .\Class_IconCatalogItem.ps1 -Force
Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\ListviewSearchLib.ps1 -Force
Import-Module -Name .\ListViewSortLib.ps1 -Force
Import-Module -Name .\PCVolumeControl.ps1 -Force
Import-Module -Name .\PropertySheetDialog.ps1 -Force
Import-Module -Name .\UtilitiesLib.ps1 -Force
Import-Module -Name .\WinFormsLibrary.ps1 -Force
#endregion

#region Type Import
Add-Type -A PresentationCore
Add-Type -A System.Drawing
Add-Type -A System.Windows.Forms
#endregion

#region Script Level Variables
$MediaPlayer = [Windows.Media.MediaPlayer]::New()
$PSCoreFont = [Drawing.Font]::New('Segoe UI',9,[Drawing.FontStyle]::Regular)
$PausePlayback = `
$StopPlayback = `
$LvwSortEnabled = $True
$PlayListErrors = $False
$PlayListDuration = [TimeSpan]0
$App = [PSCustomObject][Ordered]@{Name='PS Audio Player';Vers='Version: 1.0u - 04/04/2025'}
$AudioVolume = [PSCustomObject][Ordered]@{Min=0;Max=100}
$IconSize = [PSCustomObject][Ordered]@{Form=16;LgIco=32;Logo=64;SmIco=24;Splash=256}
$FormSize = [PSCustomObject][Ordered]@{Base=0;Min=0;Mini=0}
$AutoSize = -[Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent #Must be Negative
$PlayListHeader = '*-<{0} - Playlist Header>-*' -f $App.Name
$Disabled = '{0} Disabled, During Current Operation'
$AudioFileTypes = @('.acc','.aif','.aiff','.au','.m4a','.mp3','.snd','.wav','.wma')
$LvwColumnWidths = @($IconSize.SmIco,$AutoSize,$AutoSize)
$HelpFont = @()
$RegKeys = [Enum]::GetNames([RegistryKey])
$RegKeys[[RegistryKey]::Default] = '({0})' -f $RegKeys[[RegistryKey]::Default]
#endregion

#region Set Application Base Registry Key
Set-MasterAppName -Name $App.Name #Set Master AppName for Registry Library
$BaseKey = Get-RegistryProperty -Name $RegKeys[[RegistryKey]::Default]
if($BaseKey -ne $App.Vers.Substring(9)){
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::Default] -Value $App.Vers.Substring(9) -Type String}
#endregion

#region Get LockSettings value from Registry
# 3 Possible values: $Null(Non-Existent),0(Unlocked),1(Locked)
$LockSet = Get-RegistryProperty -Name 'LockSettings'
if($Null -eq $LockSet){$LockSet = 0}
#if $LockSettings set by user set new value
if($LockSettings.Length -gt 0){
	$LockSet = ConvertTo-Binary -Value $LockSettings
	Set-RegistryProperty -Name 'LockSettings' -Value $LockSet -Type Binary}
#endregion

#region Utility functions
function New-FileObject{
	param([Parameter(Mandatory)][String]$Path)
	$Name = [IO.Path]::GetFileName($Path)
	$Dir = [IO.Path]::GetDirectoryName($Path)
	$IsDir = Test-Path -LiteralPath $Path -PathType Container
	if(!$IsDir){
		$NameNoExt = [IO.Path]::GetFileNameWithoutExtension($Name)
		$Ext = [IO.Path]::GetExtension($Name)}
	else{
		$Name = `
		$NameNoExt = `
		$Ext = $Null}
	[PSCustomObject][Ordered]@{
		FullPath = [IO.Path]::GetFullPath($Path)
		Directory = [PSCustomObject][Ordered]@{
			Exists = [IO.Directory]::Exists($Dir)
			Name = $Dir}
		File = [PSCustomObject][Ordered]@{
			Exists = [IO.File]::Exists($Path)
			Name = $Name
			NameWithoutExtension = $NameNoExt
			Extension = $Ext} 
		IsDirectory = $IsDir}
}

function Get-ShortcutKey{
	param(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet(
			'Open Playlist','New Playlist','Edit Playlist','Reload Playlist','Exit','Find','Find Next','Font Settings',
			'Help','About','Lock Volume','Save Settings','Delete Settings','Host Information','Reset Column Width','Info')]
		[String]$Mode)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$Modes = $MyParam['Mode'].Attributes.ValidValues
	$WFK = [Windows.Forms.Keys]

	switch([Array]::IndexOf($Modes,$Mode)){
		 0 {$WFK::Alt -bor $WFK::O}
		 1 {$WFK::Alt -bor $WFK::N}
		 2 {$WFK::Alt -bor $WFK::E}
		 3 {$WFK::F5}
		 4 {$WFK::Alt -bor $WFK::F4}
		 5 {$WFK::Alt -bor $WFK::F}
		 6 {$WFK::F3}
		 7 {$WFK::Control -bor $WFK::F}
		 8 {$WFK::F1}
		 9 {$WFK::Shift -bor $WFK::F1}
		10 {$WFK::Alt -bor $WFK::L}
		11 {$WFK::Control -bor $WFK::S}
		12 {$WFK::Control -bor $WFK::Alt -bor $WFK::D}
		13 {$WFK::Control -bor $WFK::F1}
		14 {$WFK::Control -bor $WFK::R}
		15 {$WFK::Alt -bor $WFK::I}
		default {$WFK::None}
	}
}

function Invoke-Notepad{
	param([Parameter()][ValidateNotNullOrEmpty()][ValidateSet('New','Edit')][String]$Mode='New')
	if($Mode -eq 'New'){
		$F = Join-Path -Path $Env:Temp -ChildPath 'New.apl'
		$PlayListHeader|Set-Content -Path $F
		Notepad.exe $F
		Start-Sleep -Seconds 1
		Remove-Item -Path $F}
	else{Notepad.exe $PlayList}
}

function Set-ButtonEnabledState{
	param(
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('PlayListLoading','PlayListLoaded','PlayClicked','PauseClicked','StopClicked','AllOff')]
		[String]$Mode = 'AllOff')

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$Modes = $MyParam['Mode'].Attributes.ValidValues

	#region ScriptBlocks
	$CommonSB = {
		param([Boolean]$P1,[Boolean]$P2,[Boolean]$P3)
		$Buttons[[MediaButton]::Play].Enabled = $P1
		$Buttons[[MediaButton]::Pause].Enabled = $P2
		$Buttons[[MediaButton]::Stop].Enabled = $P2
		0..0+3..$FileMenuItems.GetUpperBound(0)|ForEach-Object{$FileMenuItems[$_].Enabled = $P3}
		0..0+3..5+$LvwCtxMenuItems.GetUpperBound(0)|ForEach-Object{$LvwCtxMenuItems[$_].Enabled = $P3}
	}
	$PlayListLoading = {
		0..0+3..$FileMenuItems.GetUpperBound(0)|ForEach-Object{$FileMenuItems[$_].Enabled = !$FileMenuItems[$_].Enabled}
		0..0+3..5+$LvwCtxMenuItems.GetUpperBound(0)|ForEach-Object{$LvwCtxMenuItems[$_].Enabled = !$LvwCtxMenuItems[$_].Enabled}
	}
	$PauseClicked = {
		$Buttons[[MediaButton]::Play].Enabled = $False
		$Buttons[[MediaButton]::Stop].Enabled = !$Buttons[[MediaButton]::Stop].Enabled
	}
	#endregion

	switch([Array]::IndexOf($Modes,$Mode)){
		0 {Invoke-Command -ScriptBlock $PlayListLoading}
		1 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList (!$PlayListErrors),$False,$True}
		2 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList $False,$True,$False}
		3 {Invoke-Command -ScriptBlock $PauseClicked}
		4 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList $True,$False,$True}
		Default {foreach($Button in $Buttons){$Button.Enabled = $False}}}
}

function Set-ListViewColumnWidths{
	for($C=0;$C -lt $ListView1.Columns.Count;$C++){
		$ListView1.Columns[$C].Width=$LvwColumnWidths[$C]
	}
}

function Set-MenuItem{
Param(
	[Parameter(Mandatory)][String[]]$Labels,
	[Parameter(Mandatory)][Windows.Forms.ToolStripMenuItem[]]$MenuItems,
	[Parameter(Mandatory)][Drawing.Size]$ItemSize,
	[Parameter(Mandatory)][String]$ItemPrefix,
	[Parameter()][Array]$HotKeys = @(),
	[Parameter()][Switch]$SetSizeOff,
	[Parameter()][Switch]$NoHotKeys)

for($C=0;$C -le $MenuItems.GetUpperBound(0);$C++){
	$MenuItems[$C].Name =  $ItemPrefix + ($C+1)
	$MenuItems[$C].Text = $Labels[$C]
	if(!$SetSizeOff){$MenuItems[$C].Size = $ItemSize}
	if(!$NoHotKeys){
		$MenuItems[$C].ShortcutKeys =`
		Get-ShortcutKey -Mode $(if($Hotkeys.Length -eq 0) {$Labels[$C]} else {$HotKeys[$C]})
		}
	}
}

#Gets Media File Duration String using Shell.ExtendedProperty Method
function Get-MediaDuration([String]$Path){
	$DurationMask = '{0:hh\:mm\:ss\.fff}' #Hours:Minutes:Seconds.Milliseconds
	$Folder = Split-Path -Parent -Path $Path
	$File = Split-Path -Leaf -Path $Path
	$Shell = New-Object -COMObject Shell.Application
	$ShellFolder = $Shell.NameSpace($Folder)
	$ShellFile = $ShellFolder.ParseName($File)
	$t = [uint64]$ShellFile.ExtendedProperty("System.Media.Duration") #Media File Duration in Ticks (100ns units)
	return $DurationMask -f [TimeSpan]::FromTicks([Long]$t) #Media File Duration String
}

function Add-ListViewItem{
	param(
		[Parameter(Mandatory)][Alias('Lvw')][Windows.Forms.ListView]$Control,
		[Parameter(Mandatory)][Alias('Value')][String]$MainValue,
		[Parameter(Mandatory)][Alias('Idx')][Int]$IconIndex)

	$RunTime = Get-MediaDuration -Path $MainValue
	$Script:PlayListDuration += [TimeSpan]$RunTime
	$LvwItem = [Windows.Forms.ListViewItem]::New($MainValue,$IconIndex)
	[Void]$LvwItem.SubItems.Add($RunTime)
	[Void]$LvwItem.SubItems.Add([IO.Path]::GetFileName($MainValue)) 
	[Void]$Control.Items.Add($LvwItem)
	$LblStatus.Text = 'Opening Playlist (File: {0}), Please Wait ...' -f $Control.Items.Count
	$LblTotalRuntime.Text = 'Playlist Duration: [{0:dd\:hh\:mm\:ss\.fff}] - Files: {1}' -f $Script:PlayListDuration,$Control.Items.Count
	[Windows.Forms.Application]::DoEvents()
}

function Set-FocusedListViewItem{
$ListView1.Items[0].Selected = `
$ListView1.Items[0].Focused = $True
$ListView1.Items[0].EnsureVisible()
}

function Confirm-Playlist{
	param([Parameter(Mandatory)][Alias('Items')][Array]$PlayListItems)

	$HdrValid = ($PlayListItems[0] -eq $PlayListHeader)
	$ErrorList = @()
	if($HdrValid){
		$ArrayList = [Collections.ArrayList]::New()
		$ArrayList.AddRange($PlayListItems)
		$ArrayList.RemoveAt(0) #Remove Playlist Header
		$LineNo = 2 #1-Based Line# in Playlist
		foreach($Item in $ArrayList){
			if($Item.Length -eq 0 -or $Item.StartsWith('*')){$LineNo++;Continue}
			$FO = New-FileObject -Path $Item
			if(!$(if(!$FO.IsDirectory){$FO.File.Exists}else{$FO.Directory.Exists}) -or
				$FO.File.Exists -and !$AudioFileTypes.Contains($FO.File.Extension)){$ErrorList += $LineNo}
			$LineNo++
		}
	}
	[PSCustomObject][Ordered]@{
		HdrIsValid = $HdrValid
		LineErrors = $ErrorList.Length -gt 0
		ArrayList = $ArrayList
		Errors = $ErrorList} 
}

function Remove-Comments{
	param([Parameter(Mandatory)][Collections.ArrayList]$Items)

	$Comments = @()
	for($C = 0;$C -lt $Items.Count;$C++){
		if($Items[$C].Length -eq 0 -or $Items[$C].StartsWith('*')){$Comments += $C}
	}
	if($Comments.Length -gt 0){
		[Array]::Reverse($Comments)
		$Comments|ForEach-Object{$Items.RemoveAt($_)}
	}
	return $Items
}

function Show-PlaylistErrors{
	param([Parameter(Mandatory)][Alias('V')][PSCustomObject]$Validation)

	$Script:PlayListErrors = $True
	$HdrErr = 'Invalid or Missing Playlist Header'
	if($AutoPlay.IsPresent){
		#Output Error Log Header
		$App.Name+" Errors {0}`r`nPlaylist: {1}`r`n" -f (Get-Date),$PlayList|Out-File -FilePath $ErrorLog
		if($Validation.HdrIsValid)
			{#Report Line Errors
			foreach($Line in $Validation.Errors)
				{"Line#: {0}" -f $Line|Out-File -FilePath $ErrorLog -Append}
			}
		else{"$HdrErr`r`n"|Out-File -FilePath $ErrorLog -Append}
		}
	else{
		$E = 'Playlist Errors Detected'
		if($Validation.HdrIsValid){
			$ErrMsg = "{0} at Lines:`r`n{1}" -f $E,(Convert-ArrayToString -Array $Validation.Errors)}
		else{$ErrMsg = "{0}:`r`n{1}" -f $E,$HdrErr}
		Show-MessageBoxEx -M $ErrMsg -T $Form1.Text -B OK -I Warning -D Button1
		}
}

function Install-ValidPlaylist{
$Script:PlayListDuration = [TimeSpan]0
$ListView1.Items.Clear()
$ListView1.BeginUpdate()
foreach($Line in $RV.ArrayList){
	$FO = New-FileObject -Path $Line 
	if($FO.IsDirectory)
		{#Expand Directory to Child Files
		Get-ChildItem -LiteralPath $Line -Recurse:$Recurse|Sort-Object -Property FullName| `
		ForEach-Object{
			if($AudioFileTypes.Contains($_.Extension))
				{Add-ListViewItem -Lvw $ListView1 -Value $_.FullName -Idx ([MediaIcon]::AudioFile)}}
		}
	else{Add-ListViewItem -Lvw $ListView1 -Value $Line -Idx ([MediaIcon]::AudioFile)}
	}
if($ListView1.Items.Count -gt 0){
	$FileMenuItems[[FileMenuItem]::ReloadPlaylist].Enabled = `
	$LvwCtxMenuItems[[LvwCtxItem]::ReloadPlaylist].Enabled = $True
	Set-ListViewColumnWidths
	}
$ListView1.EndUpdate()
}

function Open-Playlist{
param([Parameter()][Alias('NP')][Switch]$NoPrompt)

if($LvwSortEnabled -eq $False){
	Show-MessageBox -M ($Disabled -f 'Open') -T $Form1.Text -Icon Warning
	}
else{
	$LvwSortEnabled = Toggle-Boolean -Target $LvwSortEnabled #Disable

	if(!$NoPrompt.IsPresent){ 
		$RV = Invoke-OpenFileDialog `
			-ReturnMode Filename `
			-FileFilter ($App.Name+' Playlist (*.apl)|*.apl|Text files (*.txt)|*.txt|All files (*.*)|*.*')`
			-Title $App.Name}
	else{$RV = $PlayList}

	if($RV.length -gt 0){
		$Script:PlayList = $RV
		$Script:PlayListErrors = $False
		if(Test-Exists -Mode File -Location $Script:PlayList){
			Set-ButtonEnabledState -Mode PlayListLoading
			$RV = Confirm-Playlist -Items (Get-Content -Path $Script:PlayList)
			if($Null -ne $RV.ArrayList){$RV.ArrayList = Remove-Comments -Items $RV.ArrayList}
			if($RV.HdrIsValid -and !$RV.LineErrors){Install-ValidPlaylist}
			else{Show-PlaylistErrors -V $RV}
			}
		else{
			$Msg = 'Playlist:  {0},  Not Found!' -f [IO.Path]::GetFileName($Script:PlayList)
			Show-MessageBoxEx -O $Form1 -M $Msg -T $Form1.Text -Icon Warning
			}
		}
	$LblStatus.Text = "Total Files in Playlist (After Expansion):  {0}" -f $ListView1.Items.Count
	if($ListView1.Items.Count -gt 0){
		Set-FocusedListViewItem
		Set-ButtonEnabledState -Mode PlayListLoaded}
	else{Set-ButtonEnabledState -Mode PlayListLoading}
	$ListView1.Columns[[LvwColumn]::File].Width -= ($ListView1.Columns[[LvwColumn]::File].Width*.03)
	$LvwSortEnabled = Toggle-Boolean -Target $LvwSortEnabled #Enable
	}
}

function Invoke-AudioFile{
	param(
		[Parameter(Mandatory,ValueFromPipeline)][Alias('P')][uri]$Path,
		[Parameter()][Alias('D')][Single]$OpenDelay = 2.75)

	$TimeFormat = 'hh\:mm\:ss\.fff'	
	$Get_Position = {"Position:  [{0:$TimeFormat}]" -f $MediaPlayer.Position}
	$MediaPlayer.Open($Path)
	Start-Sleep -Milliseconds ($OpenDelay*1000) #This allows the player time to load the audio file
	$MediaPlayer.Volume = 1
	$LblRuntime.Text = "Duration: [{0:$TimeFormat}]" -f $MediaPlayer.NaturalDuration.TimeSpan
	$MediaPlayer.Play()
	Do	{
		$LblPosition.Text = & $Get_Position
		if(([Audio]::Volume*100) -ne $Slider.Value){
			if($ToolMenuItems[[ToolMenuItem]::LockVolume].Checked){
				[Audio]::Volume = $Slider.Value/100} #Enforce Volume Lock
			else{
				$Slider.Value = [Int]([Audio]::Volume*100)} #Update Slider Value
		}
		[Windows.Forms.Application]::DoEvents()
		if($PausePlayback -eq $True){$MediaPlayer.Pause()}
	}
	Until($MediaPlayer.Position -eq $MediaPlayer.NaturalDuration.TimeSpan -or $StopPlayback -eq $True)
	$LblPosition.Text = & $Get_Position
	$MediaPlayer.Stop()
	$MediaPlayer.Close()
}

function Invoke-Playlist{
if($LvwSortEnabled -eq $False){
	Show-MessageBox -M ($Disabled -f 'Play') -T $Form1.Text -Icon Warning}
else{
	$LvwSortEnabled = Toggle-Boolean -Target $LvwSortEnabled #Disable
	$Script:PausePlayback = `
	$Script:StopPlayback = $False

	for($C=$ListView1.SelectedItems[0].Index;$C -lt $ListView1.Items.Count;$C++){
		$ListView1.Items[$C].Selected = `
		$ListView1.Items[$C].Focused = $True
		$ListView1.Items[$C].EnsureVisible()
		$LblStatus.Text = 'Now Playing:  {0}' -f $ListView1.SelectedItems[0].Subitems[[LvwColumn]::File].Text
		Invoke-AudioFile -Path $ListView1.Items[$C].Text
		if($StopPlayback -eq $True){$C=$ListView1.Items.Count}
		#Loopback Control
		if($CheckBoxes[[CheckboxID]::Loop].Checked -and $C -eq $ListView1.Items.Count-1){
			if($AutoClose){
				Invoke-Command -ScriptBlock $Exit_Click
				break}
			$C = -1}
		}
	if($AutoClose)
		{Invoke-Command -ScriptBlock $Exit_Click}
	else
		{Set-ButtonEnabledState -Mode StopClicked}
	$LvwSortEnabled = Toggle-Boolean -Target $LvwSortEnabled #Enable
	}
}

function Save-Settings{
	if($LockSet -eq 1){
		Show-MessageBox -M 'Registry Settings Locked by Admin' -T $App.Name -B OK -I Warning
		return
	}

	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::AutoClose]`
		-Value $CheckBoxes[[CheckboxID]::AutoClose].Checked -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::AutoPlay]`
		-Value $AutoPlay.IsPresent -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::HelpRtbFont]`
		-Value $HelpFont -Type MultiString
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::HideVolumeLock]`
		-Value $HideLockVolume.IsPresent -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::LockVolume]`
		-Value $ToolMenuItems[[ToolMenuItem]::LockVolume].Checked -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::LoopPlayback]`
		-Value $CheckBoxes[[CheckboxID]::Loop].Checked -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::MainFormSize]`
		-Value @($Form1.Size.Width,$Form1.Size.Height) -Type MultiString
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::MainLvwColumnWidth]`
		-Value @($ListView1.Columns.Width) -Type MultiString
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::MainLvwFont]`
		-Value @($ListView1.Font.Name,$ListView1.Font.Size,$ListView1.Font.Style) -Type MultiString
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::Minimized]`
		-Value $Minimized.IsPresent -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::MiniMode]`
		-Value $ChkMini.Checked -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::Playlist]`
		-Value $PlayList -Type String
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::RecurseDirectory]`
		-Value $CheckBoxes[[CheckboxID]::Recurse].Checked -Type Binary
	Set-RegistryProperty -Name $RegKeys[[RegistryKey]::Volume]`
		-Value $Slider.Value -Type DWord

	$ToolMenuItems[[ToolMenuItem]::DeleteSettings].Enabled = $True
}

function Get-Settings{
	$NewObj = [PSObject]::New()
	for($C=0;$C -lt $RegKeys.Count;$C++){
		Add-Member -InputObject $NewObj -MemberType NoteProperty -Name $RegKeys[$C]`
			-Value (Get-RegistryProperty -Name $RegKeys[$C])
	}
	return $NewObj
}

function Install-Settings{
	param([Parameter(Mandatory)][PSObject]$RS)

	$B2S = {param($B);[Switch]($B -eq 1)}

	if($Script:PlayList.Length -eq 0 -and $RS.Playlist.Length -gt 0){$Script:PlayList=$RS.Playlist}
	if(!$AutoPlay.IsPresent){$Script:AutoPlay = & $B2S $RS.AutoPlay}
	if(!$AutoClose.IsPresent){$Script:AutoClose = & $B2S $RS.AutoClose}
	if(!$Recurse.IsPresent){$Script:Recurse = & $B2S $RS.RecurseDirectory}
	if(!$LoopPlayback.IsPresent){$Script:LoopPlayback = & $B2S $RS.LoopPlayback}
	if(!$LockVolume.IsPresent){$Script:LockVolume = & $B2S $RS.LockVolume}
	if(!$HideLockVolume.IsPresent){$Script:HideLockVolume = & $B2S $RS.HideVolumeLock}
	if(!$MiniMode.IsPresent){$Script:MiniMode = & $B2S $RS.MiniMode}
	if(!$Minimized.IsPresent){$Script:Minimized = & $B2S $RS.Minimized}
	if($Volume -eq -1 -and $RS.Volume -gt 0){$Script:Volume = $RS.Volume}
	if($RS.MainFormSize.Length -gt 0){
		$RS.MainFormSize = [Drawing.Size]::New($RS.MainFormSize[0],$RS.MainFormSize[1])
	}
	if($RS.HelpRtbFont.Length -gt 0){
		$Script:HelpFont = $RS.HelpRtbFont
		$RS.HelpRtbFont = Build-FontObject -Font $RS.HelpRtbFont
	}
	if($RS.MainLvwFont.Length -gt 0){
		$RS.MainLvwFont = Build-FontObject -Font $RS.MainLvwFont
	}

	return $RS
}

function Build-FontObject{
	param([Parameter(Mandatory)][String[]]$Font)
	if($Font[2] -eq 'Bold, Italic'){$Font[2] = 'BoldItalic'}
	[PSCustomObject][Ordered]@{Name = $Font[0];Size = $Font[1];Style = $Font[2]}
}
#endregion Utility functions

#region Add Custom DLL
Load-DLL -DLL 'BlueflameDynamics.IconTools.dll' -Msg
#endregion

#region Initalization Routines
if($Playlist.StartsWith('.\')){$Playlist = Resolve-CurrentLocation -Path $Playlist}
$FindObj = [LvwSearchValueItem]::New()
$RegistrySettings = Install-Settings -RS (Get-Settings)

#region Build Icon DLL Catalog
<#	Array items defined in the order of the icons within the DLL
	The ControlIndex value of -1 identifies an unused icon, while
	other values set the desired sort order for importing the icons
	into 1 or more imagelists.
#>
$IconCatalog = @(
('App Icon','Open Playlist','New Playlist','Edit Playlist','Reload Playlist','Directory',
'Audio File','Find','Find Next','Font Settings','Help','Exit','Info','Play','Pause','Stop','Logo'),
(-1,0,1,2,3,-1,-1,4,5,6,-1,8,7,-1,-1,-1,-1)
)
$IconCatalogItems = [IconCatalogItem]::ConvertFrom2dArray($IconCatalog)
$IconCatalogItems = [IconCatalogItem]::GetSelectedItems($IconCatalogItems)
#endregion
#endregion

function Show-MainForm{
	#region Form Objects
	$Form1 = [Windows.Forms.Form]::New()
	$MainMenu = [Windows.Forms.MenuStrip]::New()
	$LvwCtxMenuStrip = [Windows.Forms.ContextMenuStrip]::New()
	$ListView1 = [Windows.Forms.ListView]::New()
	$Slider = [Windows.Forms.TrackBar]::New()
	$Panel1 = [Windows.Forms.Panel]::New()
	$PicIcon = [Windows.Forms.PictureBox]::New()
	$ChkMini = [Windows.Forms.Checkbox]::New()
	$InitialFormWindowState = [Windows.Forms.FormWindowState]::Normal
	# Control Arrays
	$CheckBoxes = New-ObjectArray -TypeName Windows.Forms.Checkbox -Count 3
	$MainMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 3
	$FileMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 7
	$ToolMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 5
	$HelpMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 3
	$LvwCtxMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 9
	$LvwCtxMenuBars = New-ObjectArray -TypeName Windows.Forms.ToolStripSeparator -Count 4
	#endregion Form Objects

	#region ImageList for nodes
	$ImageList = New-ObjectArray -TypeName Windows.Forms.ImageList -Count 2
	for ($C = 0; $C -lt $ImageList.Count; $C++){
		$ImageList[$C].ColorDepth = [Windows.Forms.ColorDepth]::Depth32Bit}
	$ImageList[[ImageListID]::SmIcon].ImageSize = [Drawing.Size]::New($IconSize.SmIco,$IconSize.SmIco)
	$ImageList[[ImageListID]::LgIcon].ImageSize = [Drawing.Size]::New($IconSize.LgIco,$IconSize.LgIco)
	#endregion

	#region Custom Code for events.
	$About_Click = {
		$AboutText = -join(
		$App.Vers,"{1}",
		'Created by Randy Turner - mailto:turner.randy21@yahoo.com',"{2}",
		'PS Audio Player was designed as a specialized task launcher.',"{2}",
		'Script Name: AudioPlayer.ps1',"{2}",   
		'Synopsis:',"{2}",
		'This script launches audio files from a defineable location by use of a WinForm',"{1}",
		"{0}","{1}",
		"Supported Audio File Types: {3}","{1}",
		"{0}","{2}",
		'For additional help run the Powershell Get-Help cmdlet.' `
		-f $('=' * 66),"`n",("`n"*2),(Convert-ArrayToString $AudioFileTypes))
		Show-AboutForm -AppName $App.Name -AboutText $AboutText -URL -FormHeight 255
	}

	$Help_Click = {
	$HelpFile = Resolve-CurrentLocation -Path '.\PS_Audio_Player_Help.txt'
	if($PausePlayback -or $StopPlayback){
		Show-HelpForm -AppName $App.Name -HelpText $HelpFile -File -Read}
	else{
		Show-HelpForm -AppName $App.Name -HelpText $HelpFile -File}
	}

	$Exit_Click = {$Form1.Close()}

	$FontSettings_Click = {Invoke-FontDialog -Control $ListView1 -FontMustExist -AllowSimulations -AllowVectorFonts}

	$Find_Click = {
		if(($RV = Show-InputBox -Prompt 'Find?' -Title 'Search') -ne ''){
			$FindObj.Initialized = $False
			Find-ListViewItem -SVI ([Ref]$FindObj) -Lvw $ListView1 -Val $RV -Row 0 -Col 2}  
	}

	$FindNext_Click = {
		if($FindObj.Initialized -eq $True)
			{Find-ListViewItem -SVI ([Ref]$FindObj) -Lvw $ListView1}
		else
			{Invoke-Command -ScriptBlock $Find_Click}
	}

	$LockVolume_Click = {
		$Idx = [ToolMenuItem]::LockVolume
		$ToolMenuItems[$Idx].Text = '{0}ock Volume' -f $(if($ToolMenuItems[$Idx].Checked){'L'}else{'Unl'})
		$Slider.Enabled = !$Slider.Enabled
		$ToolMenuItems[$Idx].Checked = (!$ToolMenuItems[$Idx].Checked)
	}

	$DeleteSettings_Click = {
		if($LockSet -eq 1){
			Show-MessageBox -M 'Registry Settings Locked by Admim' -T $App.Name -B OK -I Warning
			return
		}
		if((Show-MessageBox -M 'Confirm Delete Settings?' -T $App.Name -B YesNo -I Warning -D Button2)`
			-eq [Windows.Forms.DialogResult]::Yes){
				Flush-AppRegistryKey
				$This.Enabled = $False}
	}

	$ChkMini_Changed = {
		$MiniMode = !$MiniMode
		if($ChkMini.Checked){
			$FormSize.Base = $Form1.Size
			$Form1.Size = $Form1.MinimumSize = $FormSize.Mini}
		else{
			$Form1.MinimumSize = $FormSize.Min
			$Form1.Size = $FormSize.Base}
		RepositionTo-CenterScreen -Form $Form1
	}

	$Host_Click = {Show-HostInfo}

	$Form_BringToTop = {
		$This.TopMost = $True
		$This.BringToFront()
		$This.TopMost = $False
	}

	$Properties_Click = {
		Show-PropertySheetDialog -P $ListView1.SelectedItems[0].Text
	}
	#endregion

	#region Common Control Variables
	$DLL =  Resolve-CurrentLocation -Path '.\AudioPlayerIcons.dll'
	<#Error handler for missing/invalid DLL#>
	if((Test-Exists -Mode File -Location $DLL) -eq $False){
		$ErrMsg = "DLL File: {0} Missing or Invalid - Job Aborted!" -f $DLL
		$RV=Show-MessageBox -M $ErrMsg -T $Form1.Text -B OK -I Error
		exit
	}
	$LvwCtxMenuItemsSize = [Drawing.Size]::New(219,32)
	#endregion

	#region Form Level Code Groups
	#region Form Code
	$Form1.Name = 'Form1'
	$Form1.Text = $App.Name
	$Form1.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
	$Form1.Icon = [BlueflameDynamics.IconTools]::ExtractIcon(
		$DLL,
		[Array]::IndexOf($IconCatalog[[IconCatalogGroup]::Tag],'App Icon'),
		$IconSize.Form)
	$Form1.Size = [Drawing.Size]::New(800,510)
	$FormSize.Mini = [Drawing.Size]::New(($Form1.Size.Width*0.8875),($Form1.Size.Height*0.4588))
 	$FormSize.Min = `
	$Form1.MinimumSize = [Drawing.Size]::New(710,410)
	$Form1.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
	if($Minimized.IsPresent){$Form1.WindowState = [Windows.Forms.FormWindowState]::Minimized}
	if($RegistrySettings.MainFormSize.Width -gt 0){
		$Form1.Size = [Drawing.Size]::New($RegistrySettings.MainFormSize.Width,$RegistrySettings.MainFormSize.Height)
	}
	$Script:FullFormSize = $Form1.Size
	#endregion

	#region Populate Imagelists
	$ILN = @('Audio File','Directory')
	for ($C = 0; $C -lt $ImageList.Count; $C++){$ImageList[$C].Images.Clear()}

	for ($C = 0; $C -lt $ILN.Count; $C++){
		$ImageList[[ImageListID]::SmIcon].Images.Add(
			$ILN[$C],
			[BlueflameDynamics.IconTools]::ExtractIcon(
				$DLL,
				[Array]::IndexOf($IconCatalog[[IconCatalogGroup]::Tag],$ILN[$C]),
				$IconSize.SmIco))
	}

	for ($C = 0; $C -le $LvwCtxMenuitems.GetUpperBound(0); $C++){
		$ImageList[[ImageListID]::LgIcon].Images.Add(
			$IconCatalogItems[$C].Tag,
			[BlueflameDynamics.IconTools]::ExtractIcon($DLL, $IconCatalogItems[$C].IconIndex, $IconSize.LgIco))
	}
	#endregion 

	#region MainMenu 
	<#
	MainMenu is a Drop-Down menu designed to provide access to the
	various functions.
	#>
	$MainMenu.Name = 'MainMenu'
	$MainMenu.Size = [Drawing.Size]::New(220,30)
	$MainMenu.Items.AddRange($MainMenuItems)
	$MainMenuItemSize = [Drawing.Size]::New(219,22)

	Set-MenuItem (Split-EnumNames -Enum ([FileMenuItem])) $FileMenuItems $MainMenuItemSize 'FileMenuItem'
	$FileMenuItems[[FileMenuItem]::OpenPlaylist].Add_Click({Open-Playlist})
	$FileMenuItems[[FileMenuItem]::NewPlaylist].Add_Click({Invoke-Notepad -Mode New})
	$FileMenuItems[[FileMenuItem]::EditPlaylist].Add_Click({Invoke-Notepad -Mode Edit})
	$FileMenuItems[[FileMenuItem]::ReloadPlaylist].Add_Click({Open-Playlist -NP})
	$FileMenuItems[[FileMenuItem]::Find].Add_Click($Find_Click)
	$FileMenuItems[[FileMenuItem]::FindNext].Add_Click($FindNext_Click)
	$FileMenuItems[[FileMenuItem]::Exit].Add_Click($Exit_Click)
	$FileMenuItems[[FileMenuItem]::ReloadPlaylist].Enabled = $False
	$FileMenuBar = New-ObjectArray -TypeName Windows.Forms.ToolStripSeparator -Count 2
	$X = $ImageList[[ImageListID]::LgIcon].Images
	$FileMenuItems|ForEach-Object{$_.Image = $X[$X.IndexOfKey($_.Text)]}
	Remove-Variable X

	Set-MenuItem (Split-EnumNames -Enum ([ToolMenuItem])) $ToolMenuItems $MainMenuItemSize 'ToolMenuItem'
	$ToolMenuItems[[ToolMenuItem]::ResetColumnWidth].Add_Click({Set-ListViewColumnWidths})
	$ToolMenuItems[[ToolMenuItem]::FontSettings].Add_Click($FontSettings_Click)
	$ToolMenuItems[[ToolMenuItem]::LockVolume].Add_Click($LockVolume_Click)
	$ToolMenuItems[[ToolMenuItem]::SaveSettings].Add_Click({Save-Settings})
	$ToolMenuItems[[ToolMenuItem]::DeleteSettings].Add_Click($DeleteSettings_Click)
	$ToolMenuItems[[ToolMenuItem]::LockVolume].Enabled = `
	$ToolMenuItems[[ToolMenuItem]::LockVolume].Visible = !$HideLockVolume.IsPresent

	Set-MenuItem (Split-EnumNames -Enum ([HelpMenuItem])) $HelpMenuItems $MainMenuItemSize 'HelpMenuItem'
	$HelpMenuItems[[HelpMenuItem]::Help].Add_Click($Help_Click)
	$HelpMenuItems[[HelpMenuItem]::About].Add_Click($About_Click)
	$HelpMenuItems[([HelpMenuItem]::HostInformation)].Add_Click($Host_Click)

	Set-MenuItem (Split-EnumNames -Enum ([MainMenuItem])) $MainMenuItems $MainMenuItemSize 'MainMenuItem' -SetSizeOff -NoHotKeys
	$MainMenuItems[[MainMenuItem]::File].DropDownItems.AddRange($FileMenuItems)
	$MainMenuItems[[MainMenuItem]::Tools].DropDownItems.AddRange($ToolMenuItems)
	$MainMenuItems[[MainMenuItem]::Help].DropDownItems.AddRange($HelpMenuItems)
	$MainMenuItems[[MainMenuItem]::File].DropDownItems.Insert([FileMenuItem]::Exit,$FileMenuBar[1])
	$MainMenuItems[[MainMenuItem]::File].DropDownItems.Insert([FileMenuItem]::Find,$FileMenuBar[0])
	$Form1.Controls.Add($MainMenu)
	#endregion MainMenu

	#region Labels
	$LblRuntime = [Windows.Forms.Label]::New()
	$LblRuntime.Name = 'LblRuntime'
	$LblRuntime.Parent = $Panel1
	$LblRuntime.Location = [Drawing.Point]::New(5,5)
	$LblRuntime.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblRuntime.BackColor = $CpColor = [Drawing.SystemColors]::Control
	$LblRuntime.Text = ''
	$LblRuntime.Font = $PSCoreFont
	$LblRuntime.Height = 18
	$LblRuntime.Width = 140
	$LblRuntime.AutoEllipsis = $True
	$LblRuntime.Anchor = Get-Anchor -T -L

	$LblPosition = [Windows.Forms.Label]::New()
	$LblPosition.Name = 'LblPosition'
	$LblPosition.Parent = $Panel1
	$LblPosition.Location = [Drawing.Point]::New(5,(5+$LblRuntime.Height+2))
	$LblPosition.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblPosition.BackColor = $CpColor
	$LblPosition.Text = ''
	$LblPosition.Font = $PSCoreFont
	$LblPosition.Height = $LblRuntime.Height
	$LblPosition.Width = $LblRuntime.Width
	$LblPosition.AutoEllipsis = $True
	$LblPosition.Anchor = Get-Anchor -T -L

	$LblStatus = [Windows.Forms.Label]::New()
	$LblStatus.Name = 'LblStatus'
	$LblStatus.Parent = $Panel1
	$LblStatus.Location = [Drawing.Point]::New((5+$LblRuntime.Width+5),5)
	$LblStatus.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblStatus.BackColor = $CpColor
	$LblStatus.Text = ''
	$LblStatus.Font = $PSCoreFont
	$LblStatus.Width = $Panel1.Width - ($LblRuntime.Width+15)
	$LblStatus.Height = $LblRuntime.Height
	$LblStatus.AutoEllipsis = $True
	$LblStatus.Anchor = Get-Anchor -T -L -R
	
	$LblVolume = [Windows.Forms.Label]::New()
	$LblVolume.Name = 'LblVolume'
	$LblVolume.Parent = $Panel1
	$LblVolume.Font = [Drawing.Font]::new('Segoe UI',22,([Drawing.FontStyle]::Regular))
	$LblVolume.Location = [Drawing.Point]::New(160,40)
	$LblVolume.BorderStyle = [Windows.Forms.BorderStyle]::None
	$LblVolume.Text = 'Vol: {0}' -f [Int]([Audio]::Volume*100)
	$LblVolume.Width = 120
	$LblVolume.Height = 45
	$LblVolume.AutoEllipsis = $True
	$LblVolume.Anchor = Get-Anchor -T -L

	$LblTotalRuntime = [Windows.Forms.Label]::New()
	$LblTotalRuntime.Name = 'LblTotRuntime'
	$LblTotalRuntime.Parent = $Panel1
	$LblTotalRuntime.Location = [Drawing.Point]::New(281,82)
	$LblTotalRuntime.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblTotalRuntime.BackColor = $CpColor
	$LblTotalRuntime.Text = ''
	$LblTotalRuntime.Font = $PSCoreFont
	$LblTotalRuntime.Height = $LblRuntime.Height
	$LblTotalRuntime.Width = 300
	$LblTotalRuntime.AutoEllipsis = $True
	$LblTotalRuntime.Anchor = Get-Anchor -T -L 
	#endregion 

	#region Checkboxes
	$CBT = @('Loop Playback','Auto Close','Recurse Directories Opening Playlist')
	for($C=0;$C -lt $CheckBoxes.Count;$C++){
		$CheckBoxes[$C].Name = 'Chk'+[Enum]::GetName([CheckboxID],$C)
		$CheckBoxes[$C].Parent = $Panel1
		$CheckBoxes[$C].Anchor = Get-Anchor -T -L
		$CheckBoxes[$C].Width = 250
		$CheckBoxes[$C].Location = [Drawing.Point]::New(5,$($LblPosition.Bottom+(($CheckBoxes[$C].Height*$C)*.8)))
		$CheckBoxes[$C].Text = $CBT[$C]
	}
	$CheckBoxes[[CheckboxID]::Loop].Checked = $LoopPlayback
	$CheckBoxes[[CheckboxID]::AutoClose].Checked = $AutoClose
	$CheckBoxes[[CheckboxID]::Recurse].Checked = $Recurse
	$CheckBoxes[[CheckboxID]::Loop].Add_CheckedChanged({$Script:LoopPlayback = !$Script:LoopPlayback})
	$CheckBoxes[[CheckboxID]::AutoClose].Add_CheckedChanged({$Script:AutoClose = !$Script:AutoClose})
	$CheckBoxes[[CheckboxID]::Recurse].Add_CheckedChanged({$Script:Recurse = !$Script:Recurse})
	#endregion

	#region LvwCtxMenuStrip
	$LvwCtxMenuStrip.Name = 'LvwCtxMenuStrip'
	$LvwCtxMenuStrip.Size = [Drawing.Size]::New(220,80)
	$LvwCtxMenuStrip.Items.AddRange($LvwCtxMenuItems)
	#endregion

	#region LvwCtxMenuItems
	for($C = 0; $C -le $LvwCtxMenuItems.GetUpperBound(0); $C++){
		$LvwCtxMenuItems[$C].Name = 'LvwCtxMenu' + [Enum]::GetName([LvwCtxItem],$C)
		$LvwCtxMenuItems[$C].Text = $IconCatalogItems[$C].Tag
		$LvwCtxMenuItems[$C].Size = $LvwCtxMenuItemsSize
		$LvwCtxMenuItems[$C].Image = $ImageList[[ImageListID]::LgIcon].Images[$C]
		$LvwCtxMenuItems[$C].ImageAlign = [Drawing.ContentAlignment]::MiddleLeft
		$LvwCtxMenuItems[$C].ShowShortcutKeys = $True
		$LvwCtxMenuItems[$C].ShortcutKeys = Get-ShortcutKey -Mode $LvwCtxMenuItems[$C].Text
	}
	$LvwCtxMenuItems[[LvwCtxItem]::Properties].Text = 'Properties'
	$LvwCtxMenuItems[[LvwCtxItem]::OpenPlaylist].Add_Click({Open-Playlist})
	$LvwCtxMenuItems[[LvwCtxItem]::NewPlaylist].Add_Click({Invoke-Notepad -Mode New})
	$LvwCtxMenuItems[[LvwCtxItem]::EditPlaylist].Add_Click({Invoke-Notepad -Mode Edit})
	$LvwCtxMenuItems[[LvwCtxItem]::ReloadPlaylist].Add_Click({Open-Playlist -NP})
	$LvwCtxMenuItems[[LvwCtxItem]::ReloadPlaylist].Enabled = $False
	$LvwCtxMenuItems[[LvwCtxItem]::Find].Add_Click($Find_Click)
	$LvwCtxMenuItems[[LvwCtxItem]::FindNext].Add_Click($FindNext_Click)
	$LvwCtxMenuItems[[LvwCtxItem]::FontSettings].Add_Click($FontSettings_Click)
	$LvwCtxMenuItems[[LvwCtxItem]::Properties].Add_Click($Properties_Click)
	$LvwCtxMenuItems[[LvwCtxItem]::Exit].Add_Click($Exit_Click)
	#endregion 

	#region Separator(s)
	$LvwCtxMenuStrip.Items.Insert([LvwCtxItem]::Exit, $LvwCtxMenuBars[3])
	$LvwCtxMenuStrip.Items.Insert([LvwCtxItem]::Properties, $LvwCtxMenuBars[2])
	$LvwCtxMenuStrip.Items.Insert([LvwCtxItem]::FontSettings, $LvwCtxMenuBars[1])
	$LvwCtxMenuStrip.Items.Insert([LvwCtxItem]::Find, $LvwCtxMenuBars[0])
	#endregion

	#region Listview1
	$LvwColumnNames = [Enum]::GetNames('LvwColumn')
	$LvwColumnNames[[LvwColumn]::Icon] = ''
	$ListView1.Name = 'ListView1'
	$ListView1.Parent = $Form1
	$ListView1.View = [Windows.Forms.View]::Details
	$ListView1.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$ListView1.MultiSelect = $False
	$ListView1.GridLines = `
	$ListView1.FullRowSelect = $True
	$ListView1.Size = [Drawing.Size]::New(($Form1.Width - 42),($Form1.Height - 235))
	$ListView1.Location = [Drawing.Point]::New(13,28)
	if($Null -ne $RegistrySettings.MainLvwFont){
		$FontName  = $RegistrySettings.MainLvwFont.Name
		$FontSize  = $RegistrySettings.MainLvwFont.Size
		$FontStyle = $RegistrySettings.MainLvwFont.Style
	}
	$ListView1.Font = `
		[Drawing.Font]::New($FontName,$FontSize,(Get-SelectedFontStyle -FS $FontStyle))
	$ListView1.SmallImageList = $ImageList[[ImageListID]::SmIcon]
	$ListView1.LargeImageList = $ImageList[[ImageListID]::LgIcon]
	$ListView1.ContextMenuStrip = $LvwCtxMenuStrip
	$ListView1.Anchor = Get-Anchor -T -L -B -R

	for($C=0;$C -lt $LvwColumnNames.Count;$C++){
		$ListView1.Columns.Add($LvwColumnNames[$C])|Out-Null
		$ListView1.Columns[$C].Width=$LvwColumnWidths[$C]
	}

	$ColumnClick = {
		if($LvwSortEnabled -and $This.items.Count -gt 0){
			Sort-ListView -LvwControl $This -Column $_.Column
			Set-FocusedListViewItem
		}
	}
	$ListView1.Add_ColumnClick($ColumnClick)
	if($Null -ne $RegistrySettings.MainLvwColumnWidth){
		for($C=0;$C -lt $ListView1.Columns.Count;$C++){
			$ListView1.Columns[$C].Width = $RegistrySettings.MainLvwColumnWidth[$C]
		}
	}
	#endregion 

	#region Slider
	$CVol = if($Volume -eq -1){
				[Audio]::Volume*100}
			else{
				$Volume
				[Audio]::Volume=$Volume/100
				$LblVolume.Text = 'Vol: {0}' -f [Int]([Audio]::Volume*100)
				}
	$Slider.Name = 'VolumeSlider'
	$Slider.Parent = $Panel1
	$Slider.TickStyle = [Windows.Forms.TickStyle]::Both
	$Slider.AutoSize = $False
	$Slider.Width = 300
	$Slider.Height = 42
	$Slider.Minimum = $AudioVolume.Min
	$Slider.Maximum = $AudioVolume.Max
	$Slider.Value = $CVol
	$Slider.Location = [Drawing.Point]::New(($LblVolume.Right+2),38) 
	$Slider.BackColor= $CpColor
	$Slider.Anchor = Get-Anchor -T -L
	$Slider.Add_ValueChanged({
		[Audio]::Volume = $This.Value/$This.Maximum
		$LblVolume.Text = 'Vol: {0}' -f $This.Value
		})
	#endregion

	#region Panel Control
	$Panel1.Name = 'Panel1'
	$Panel1.Parent = $Form1
	$Panel1.Location = [Drawing.Point]::New($ListView1.Left,($ListView1.bottom + 10))
	$Panel1.Size = [Drawing.Size]::New($ListView1.Width,115)
	$Panel1.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$Panel1.BackColor = [Drawing.Color]::LightGray
	$Panel1.Anchor = Get-Anchor -B -L -R
	#endregion

	#region ChkMini
	$ChkMini.Name = 'ChkMini'
	$ChkMini.Parent = $Form1
	$ChkMini.Anchor = Get-Anchor -B -L
	$ChkMini.Location = [Drawing.Point]::New(20,($Panel1.Bottom + 10))
	$ChkMini.Width = 250
	$ChkMini.Text = 'Enable Mini Mode'
	$ChkMini.Enabled = $True
	$ChkMini.Checked = $MiniMode
	$ChkMini.Add_CheckedChanged($ChkMini_Changed)
	#endregion 

	#region Logo
	$PicIcon.Name = 'PicIcon'
	$PicIcon.Parent = $Panel1
	$PicIcon.BackColor = [Drawing.SystemColors]::Control
	$PicIcon.ForeColor = [Drawing.SystemColors]::ControlText
	$PicIcon.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
	$PicIcon.Size = [Drawing.Size]::New(68,68)
	$PicIcon.Location = [Drawing.Point]::New(($Panel1.Right-($PicIcon.Width+18)),($Slider.Top+3))
	$PicIcon.SizeMode = [Windows.Forms.PictureBoxSizeMode]::StretchImage
	$PicIcon.Image = 
		[BlueflameDynamics.IconTools]::ExtractIcon(
			$DLL,
			[Array]::IndexOf($IconCatalog[[IconCatalogGroup]::Tag],'Logo'),
			$IconSize.Logo)
	$PicIcon.Anchor = Get-Anchor -B -R
	#endregion

	#region Buttons
	$BtnWidth = 80
	$BTT = [Enum]::GetNames([MediaButton]) 
	$Buttons = New-ObjectArray -TypeName Windows.Forms.Button -Count $BTT.Count
	$BC = $Buttons.Count
	for($C=0;$C -lt $Buttons.Count;$C++){
		$Buttons[$C].Anchor = Get-Anchor -B -R
		$Buttons[$C].Name = 'Btn'+$BTT[$C]
		$Buttons[$C].Parent = $Form1
		$Buttons[$C].Size = [Drawing.Size]::New($BtnWidth,30)
		$Buttons[$C].Location = `
			[Drawing.Point]::New(($ListView1.Right - ($BtnWidth*$BC)),($Panel1.Bottom + 5))
		$Buttons[$C].Image = [BlueflameDynamics.IconTools]::ExtractIcon(
			$DLL,
			[Array]::IndexOf($IconCatalog[[IconCatalogGroup]::Tag],$BTT[$C]),
			$IconSize.SmIco)
		$Buttons[$C].ImageAlign = [Drawing.ContentAlignment]::MiddleLeft
		$Buttons[$C].Text = $BTT[$C].PadLeft($BTT[$C].Length+$(If($C -ne 1){1}Else{2}))
		$Buttons[$C].TextAlign = [Drawing.ContentAlignment]::MiddleCenter
		$Buttons[$C].Enabled = $False
		$Buttons[$C].UseVisualStyleBackColor = $True
		$BC--
	}
	$BtnPlay_Click = {
		Set-ButtonEnabledState -Mode PlayClicked
		Invoke-Playlist}
	$BtnStop_Click = {
		Set-ButtonEnabledState -Mode StopClicked
		$Script:StopPlayback = $True
		$LblStatus.Text = ''}
	$BtnPause_Click = {
		Set-ButtonEnabledState -Mode PauseClicked
		$Script:PausePlayback = !$Script:PausePlayback
		if($PausePlayback -eq $True)
			{$This.Text = 'Resume'.PadLeft(10)}
		else{
			$This.Text = 'Pause'.PadLeft(7)
			$MediaPlayer.Play()
			}
		}
	$Buttons[[MediaButton]::Play].Add_Click($BtnPlay_Click)
	$Buttons[[MediaButton]::Pause].Add_Click($BtnPause_Click)
	$Buttons[[MediaButton]::Stop].Add_Click($BtnStop_Click)
	#endregion
	#endregion Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $Form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$Form1.Add_Load({$This.WindowState = $InitialFormWindowState})
	$Form1.Add_Shown($Form_BringToTop)
	
	if($LockVolume){Invoke-Command -ScriptBlock $LockVolume_Click}
	if($SaveSettings.IsPresent){Save-Settings}
	if($PlayList.Length -gt 0){
		Import-Module -Name .\SplashScreen.ps1 -Force
		Show-SplashScreen `
			-AppName $App.Name `
			-Image ([BlueflameDynamics.IconTools]::ExtractIcon(
				$DLL,
				[Array]::IndexOf($IconCatalog[[IconCatalogGroup]::Tag],'App Icon'),
				$IconSize.Splash))`
			-PicBackColor ([Drawing.Color]::GhostWhite)
		Open-Playlist -NP
		Close-SplashScreen
	}
	if($MiniMode.IsPresent){Invoke-Command -ScriptBlock $ChkMini_Changed}
	if($AutoPlay.IsPresent){
		$LblStatus.Text = 'AutoPlay Engaged, Please Wait ...'
		$DelayTimer = [Windows.Forms.Timer]::New()
		$DelayTimer.Enabled = $False
		$DelayTimeSeconds = 5
		$DelayTimer.Interval = 1000*$DelayTimeSeconds
		$DelayTimer_Tick={
			$DelayTimer.Stop()
			if(!$PlayListErrors)
				{Invoke-Command -ScriptBlock $BtnPlay_Click}
			elseif($AutoClose)
				{Invoke-Command -ScriptBlock $Exit_Click}
		}
		$DelayTimer.Add_Tick($DelayTimer_Tick)
		$DelayTimer.Enabled = $True
		$DelayTimer.Start()
	}

	#Show the Form
	[Void]$Form1.ShowDialog()
	$Form1.Dispose()
}

#Call the Main function
Show-MainForm
<#
.\AudioPlayer.ps1 -Play .\Playlist.apl -Loop -Auto+ -Auto- -R -Vol 25 -Mini -M
#>