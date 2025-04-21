<#
.NOTES
-------------------------------------
Name:	MovieLauncher.ps1
Version: 8.0h - 07/18/2020
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
Revision History:
v8.0h - 07/18/2020 - Added changes for Powershell Core
v8.0g - 02/17/2019 - Added support for High Res icons
v8.0b - 03/15/2017 - Added MetadataIndexLib for Windows version indepentence
--------------------------------------------------------------------------------------------

.SYNOPSIS
This script launches video files from a definable location by use of a WinForm
it also includes the ability to download a selected video to the current user's
'My Videos' folder. A GUI progress window will be displayed during the file transfers.
Features include an optional 'List' view, the addition of a 'Find Next' function 
and file downloads are now run as Jobs allowing parallel downloads.
This version has been updated to get property index numbers independent of OS Version. 
This version has been tested on Windows 10, but should work on other versions.
--------------------------------------------------------------------------------------------
Supported Video File Types are: .avi, .flv, .mp4, .mov, .m4v, .mkv, .mpg, .mpeg, .webm, .wmv
--------------------------------------------------------------------------------------------

.DESCRIPTION

To Run the script without a console window use the included
PsRun3.exe PowerShell Launcher

PsRun3 accepts only 1 parameter - the name of the script to run and its parameters
enclosed in double quotes (") any quoted parameters should be enclosed in single 
quotes ('). You can create a shortcut with a parameter string like that below
to run the script for your location(s).

%PowerShellDevLib%\PsRun3.exe MovieLauncher.ps1 'Startup Directory'"
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
.Parameter Directory Alias: Dir
Home Directory for media. Defaults to the current user My Video folder.

.PARAMETER IconSize Alias: Ico
Icon size to display in LargeIcon view
supported values are: 32,48,64,96,128.

.PARAMETER FontName Alias: Fn
Name of the font to be used.

.PARAMETER FontSize Alias: Fz
Size of the font to be used, between 9-24 points.

.PARAMETER FontStyle Alias: Fs
Style of the font to be used: Bold, Italic, BoldItalic, Regular.

.PARAMETER SupressFileDownload Alias: NoDwnLd
Supresses the ability to download media to the Host PC.

.PARAMETER View Alias: V
Determines the ListView View mode of either: LargeIcon or List.

.EXAMPLE
PS> .\MovieLauncher.ps1 -Dir "\\MediaServer01\Public\Shared Videos\Music Video\" 
-Ico 48 -Fn "Times New Roman" -Fz 16 -Fs Bold -V List
#>

[CmdletBinding()]
param(
	[Parameter()][Alias('Dir')][String]$Directory = '',
	[Parameter()][Alias('Ico')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet(32,48,64,96,128)]
		[Int]$IconSize = 64,
	[Parameter()][Alias('Fn')][String]$FontName = "Lucida Console",
	[Parameter()][Alias('Fz')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(9,24)]
		[Int]$FontSize = 12,
	[Parameter()][Alias('Fs')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Bold','Italic','BoldItalic','Regular')]
		[String]$FontStyle = 'Regular',
	[Parameter()][Alias('V')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('LargeIcon','List')]
		[String]$View = 'LargeIcon',
	[Parameter()][Alias('NoDwnLd')][Switch]$SupressFileDownload)

#region Module Import
Import-Module -Name .\Class_IconCatalogItem.ps1 -Force
Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\GetSplitPathLib.ps1 -Force
Import-Module -Name .\Invoke-CopyFile.ps1 -Force
Import-Module -Name .\ListviewSearchLib.ps1 -Force
Import-Module -Name .\MetadataIndexLib.ps1 -Force
Import-Module -Name .\UtilitiesLib.ps1 -Force
Import-Module -Name .\WinFormsLibrary.ps1 -Force
#endregion

#region Script Level Variables
$AbEnd	= $False
$App = [PSCustomObject][Ordered]@{Name = 'PS Movie Launcher';Vers = 'Version: 8.0h - 07/18/2020'}
$ValidViews = @('LargeIcon','List')
[int]$View = [Array]::IndexOf($ValidViews,$View)
#endregion

#region Utility functions
function Get-DownLoadTarget{
	$KeyInfo = @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 'My Video')
	Get-RegistryValue -Key $KeyInfo[0] -Value $KeyInfo[1]
}

function Get-FileDetails{
	param(
		[Parameter(Mandatory)][Alias('P')][String]$Path,
		[Parameter(Mandatory)][Alias('D')][Bool]$IsDir)

	$ShellJob = `
		{
		param([String]$Folder,[String]$File,[Int]$PadSize,[String[]]$PropDescs,[Int[]]$PropIndices)
		#Create Windows Shell Object
		$Shell = New-Object -COMObject Shell.Application
		$ShellFolder = $Shell.Namespace($Folder)
		$ShellFile = $ShellFolder.ParseName($File)
		$Dtls = ""
		$C=0
		foreach($Index in $PropIndices)
			{
			$RV=$ShellFolder.GetDetailsOf($ShellFile, $Index)
			if($RV.length -eq 0)
				{
				switch($PropDescs[$C])
					{
					'Directory'	  {$RV = $Folder}
					'File Extension' {$RV = [IO.Path]::GetExtension($File)}
					}
				}
			$Dtls += if($RV.length -ne 0) {" {0}`t[{1}]`n" -f $PropDescs[$C].PadRight($PadSize, " "), $RV.Trim()}
			$C++
			}
		return $Dtls
		}
	
	if($IsDir -eq $True)
		{
		$PadSize = 19
		$Props = @('Name','Item type','Path','Date created','Date modified','Total size','Space free','Space used')
		$PropDescs = @('Directory:','File Type:','Full Path:','Date Created:',
			'Date Modified:','Total Volume Space:','Volume Free Space:','Volume Space Used:')
		}
	else
		{
		$PadSize = 20
		$Props = @('Folder name','Name','File extension','Item type','Date created','Date modified',
			'Date accessed','Title','Subtitle','Publisher','Genre','Length','Album','Contributing artists',
			'Year','Parental rating','Rating','Frame width','Frame height','Frame rate','Video orientation','Bit rate',
			'Data rate','Total bitrate','Size','Video compression')
		$PropDescs = @('Directory','Filename','File Extension','Item Type','Date Created','Date Modified',
			'Date Accessed','Title','Subtitle','Publisher','Genre','Run Time','Album','Contributing Artists',
			'Year Released','Parental Rating','Rating','Frame Width','Frame Height','Frame Rate','Video Orientation',
			'Bitrate','Data Rate','Total Bitrate','File Size','Video Compression')
		}

	# Get Property Index Numbers
	$MetaDataIndex = Get-MetadataIndex $path
	$SelectedProps = @()
	for($C=0;$C -le $Props.GetUpperBound(0);$C++)
		{$SelectedProps += Get-IndexByMetadataName $MetaDataIndex $Props[$C]}
	#Parse Path
	$Folder = Split-Path -Path $path
	$File = Split-Path -Path $path -Leaf
	$Form1.Cursor = Get-Cursor -Mode WaitCursor
	$RV = Start-Job `
			-Name 'MovieLauncherJob' `
			-ScriptBlock $ShellJob `
			-ArgumentList $Folder,$File,$PadSize,$PropDescs,$SelectedProps| `
		  Receive-Job -Wait -AutoRemoveJob
	$Form1.Cursor = Get-Cursor -Mode Default
	return $RV
}

function Get-ShortcutKey{
	param(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Open\Play','Download File to My Videos','Back','Home','Refresh','Exit',
					 'Find','Font Settings','Open Video Folder','Properties','About','Toggle View','Find Next')]
		[String]$Mode)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$ValidModes = $MyParam['Mode'].Attributes.ValidValues

	$WFK = [Windows.Forms.Keys]

	switch([Array]::IndexOf($ValidModes,$Mode))
		{
		0  {$WFK::Alt -bor $WFK::Enter}
		1  {$WFK::Alt -bor $WFK::D}
		2  {$WFK::Alt -bor $WFK::Left}
		3  {$WFK::Alt -bor $WFK::H}
		4  {$WFK::F5}
		5  {$WFK::Alt -bor $WFK::F4}
		6  {$WFK::Alt -bor $WFK::F}
		7  {$WFK::Control -bor $WFK::F}
		8  {$WFK::Alt -bor $WFK::V}
		9  {$WFK::Alt -bor $WFK::I}
		10 {$WFK::F1}
		11 {$WFK::Control -bor $WFK::V}
		12 {$WFK::F3}
		}
}

function Move-ListViewItems{
	param(
		[Parameter(Mandatory)][Windows.Forms.ListView]$LvSource,
		[Parameter(Mandatory)][Windows.Forms.ListView]$LvTarget)

	while ($LvSource.Items.Count -gt 0)
		{
		[Windows.Forms.ListViewItem]$Itm = $LvSource.Items[0];
		$LvSource.Items.Remove($Itm);
		$LvTarget.Items.Add($Itm);
		}
}

function Set-ListViewItem{
	param(
		[Parameter(Mandatory)][Alias('Lvw')][Windows.Forms.ListView]$Control,
		[Parameter(Mandatory)][Alias('Idx')][Int]$Index,
		[Parameter(Mandatory)][Alias('Cpo')][Object]$Stream)
	$Control.Items.Add($Stream.Name)
	$Control.Items[$counter[$Index]].ImageIndex = $Index
	$Control.Items[$counter[$Index]].Tag = $IconCatalogItems[$Index].Tag
	$Control.Items[$counter[$Index]].Name = $Stream.FullName
	$counter[$Index]++
	$Script:LvwUpdated = $True
}

function Show-PropertySheet{
	param(
		[Parameter(Mandatory)][String]$TargetFile,
		[Parameter(Mandatory)][Bool]$IsDir)

	Add-Type -A System.Windows.Forms
	$Form1.Cursor = Get-Cursor -Mode WaitCursor
	#Add objects for About
	$frmProps = New-Object -TypeName Windows.Forms.Form
	$rtbProps = New-Object -TypeName Windows.Forms.RichTextBox
	$InitialFormWindowState = New-Object -TypeName Windows.Forms.FormWindowState

	#Form
	$frmProps.Name = 'FormProperties'
	$frmProps.AutoScroll = $True
	$frmProps.ClientSize = New-Object -TypeName Drawing.Size -ArgumentList 855,525
	$frmProps.DataBindings.DefaultDataSourceUpdateMode = 0
	$frmProps.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle 
	$frmProps.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$frmProps.ShowInTaskbar = $False
	$frmProps.Text = -join $App.Name,' - Properties'
	$frmProps.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],'Properties'),16)
	$frmProps.MaximizeBox = `
	$frmProps.MinimizeBox = $False

	#Rich Textbox
	$rtbProps.Name = 'rtbProps'
	$rtbProps.Anchor = Get-Anchor -T -L -B -R
	$rtbProps.BackColor = [Drawing.Color]::FromArgb(255, 240, 240, 240)
	$rtbProps.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D 
	$rtbProps.DataBindings.DefaultDataSourceUpdateMode = 0
	$rtbProps.Location = New-Object -TypeName Drawing.Point -ArgumentList 13,13
	$rtbProps.ReadOnly = $True
	$rtbProps.DetectURLs = $False
	$rtbProps.Cursor = Get-Cursor -Mode Default
	$rtbProps.Size = New-Object -TypeName Drawing.Size -ArgumentList $($frmProps.Width - 30),$($frmProps.Height * .82)
	$rtbProps.TabIndex = 0
	$rtbProps.TabStop = $False
	$rtbProps.Font = New-Object -TypeName Drawing.Font -ArgumentList 'Lucida Console',12,(Get-SelectedFontStyle -FS Regular)
	$rtbProps.Text = Get-FileDetails -Path $TargetFile -IsDir $IsDir
	$frmProps.Controls.Add($rtbProps)

	#Exit Button
	$BtnExit = New-Object -TypeName Windows.Forms.Button
	$BtnExit.Name = 'BtnExit'
	$BtnExit.Anchor = Get-Anchor -T -L -B -R
	$BtnExit.Size = New-Object -TypeName Drawing.Size -ArgumentList 80,30
	$BtnExitLeft = $rtbProps.Right - ($BtnExit.Width+10)
	$BtnExit.Location = New-Object -TypeName Drawing.Point -ArgumentList $BtnExitLeft,$($frmProps.Height * .86)
	$BtnExit.Text = 'Exit'
	$BtnExit.UseVisualStyleBackColor = $True
	$BtnExit.Add_Click({$frmProps.Close()})
	$frmProps.Controls.Add($BtnExit)
	$Form1.Cursor = Get-Cursor -Mode Default

	[Void]$frmProps.ShowDialog()
}

function Refresh-ListView{
	param([string]$path)
  
	$Script:PreviousItem = (Get-Item -LiteralPath $path).PSParentPath
	$ListView[0].Items.Clear()
	$ListView[1].Items.Clear()
	$counter = @(0, 0)
	Get-ChildItem -LiteralPath $path | Sort-Object -Property Name | ForEach-Object {
		if(!$_.PSIsContainer)
			{$FnEx = Get-SplitFilename $_.Name.ToLower() -ExtensionOnly}

		if($_.PSIsContainer)
			{Set-ListViewItem -Lvw $ListView[0] -Idx 0 -Cpo $_}
		elseif($VideoFileTypes.Contains($FnEx))
			{Set-ListViewItem -Lvw $ListView[1] -Idx 1 -Cpo $_}
		else
			{$Script:LvwUpdated = $False}

		if($Script:LvwUpdated -eq $True)
			{$Script:CurrPath = $path}
	}

	Move-ListViewItems $ListView[1] $ListView[0]
	$Form1.Refresh()
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

function Get-ActiveView{
	switch($Script:View)
		{
		0 {$RV=[Windows.Forms.View]::LargeIcon;$Script:View=1}
		1 {$RV=[Windows.Forms.View]::List;$Script:View=0}
		}
	return $RV
}

function Get-Icon{
	param(
		[Parameter(Mandatory)][Alias('I')][Int]$Index,
		[Parameter(Mandatory)][Alias('S')][Int]$Size)
	[BlueflameDynamics.IconTools]::ExtractIcon($DLL, $IconCatalogItems[$Index].IconIndex, $Size)
}
#endregion Utility functions

#region Add Custom DLL - Data Type: BlueflameDynamics.IconTools
Load-DLL -DLL '.\BlueflameDynamics.IconTools.dll' -Msg
#endregion

#region Script Level Variables
$VideoFileTypes = @('.avi','.flv','.mp4','.mov','.m4v','.mkv','.mpg','.mpeg','.webm','.wmv')
$CurrPath = ''
[Switch]$LvwUpdated = $False
$FindObj = [LvwSearchValueItem]::New()
$DownLoadTarget = Get-DownLoadTarget
if($Directory.length -eq 0){
	$Directory = $DownloadTarget
	$SupressFileDownload=[Switch]$True}
$DLL = Resolve-CurrentLocation '.\PSMovieLauncherIcons.dll'
if((Exists -Mode File -Location $DLL) -eq $False)
	{
	$ErrMsg = "DLL File: {0} Missing or Invalid - Job Aborted!" -f $DLL
	$RV=Show-MessageBox -M $ErrMsg -T $Form1.Text -B OK -I Error
	exit
	}
#endregion

#region Build Icon DLL Catalog
<#	Array items defined in the order of the icons within the DLL
	The ControlIndex value of -1 identifies an unused icon, while
	other values set the desired sort order for importing the icons
	into 1 or more imagelists.
#>
$IconCatalog = @(
('Download File to My Videos','Film Clip','Directory','Font Settings','Help','Back','Home','Open\Play', 
'Find','Find Next','Properties','App Icon','Music Folder','Refresh','TV','Video File','Open Video Folder','Exit'),
(3,-1,0,10,-1,4,5,2,7,8,9,-1,-1,6,-1,1,11,12)
)
$IconCatalogItems = [IconCatalogItem]::ConvertFrom2dArray($IconCatalog)
$IconCatalogItems = [IconCatalogItem]::GetSelectedItems($IconCatalogItems)
#endregion

function Show-MainForm{
	#region Import the Assemblies
	Add-Type -A System.Windows.Forms
	Add-Type -A System.Drawing
	#endregion

	#region Form Objects
	$Form1 = New-Object -TypeName Windows.Forms.Form
	$Label1 = New-Object -TypeName Windows.Forms.Label
	$MainMenu = New-Object -TypeName Windows.Forms.MenuStrip
	$LvwCtxMenuStrip = New-Object -TypeName Windows.Forms.ContextMenuStrip
	$InitialFormWindowState = New-Object -TypeName Windows.Forms.FormWindowState
	# Control Arrays
	$ListView = New-ObjectArray -TypeName Windows.Forms.ListView -Count 2
	$MainMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 4
	$FileMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 6
	$NaviMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 3
	$ToolMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 3
	$HelpMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 2
	$LvwCtxMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 11
	$LvwCtxMenuBars = New-ObjectArray -TypeName Windows.Forms.ToolStripSeparator -Count 2
	#endregion Form Objects

	#region ImageList for nodes
	$ImageList = New-ObjectArray -TypeName Windows.Forms.ImageList -Count 3
	for ($C = 0; $C -lt $ImageList.Count; $C++){
		$ImageList[$C].ColorDepth = [Windows.Forms.ColorDepth]::Depth32Bit}
	$ImageList[0].ImageSize = New-Object -TypeName Drawing.Size -ArgumentList $IconSize,$IconSize
	$ImageList[1].ImageSize = New-Object -TypeName Drawing.Size -ArgumentList 32,32
	$ImageList[2].ImageSize = New-Object -TypeName Drawing.Size -ArgumentList 24,24
	#endregion

	#region Custom Code for events.
	$LvwDouble_Click = {
		if($ListView[0].SelectedItems[0].Tag -eq $IconCatalogItems[1].Tag)
			{Invoke-Item -Path $ListView[0].SelectedItems[0].Name}
		elseif($ListView[0].SelectedItems[0].Tag -eq $IconCatalogItems[0].Tag)
			{Refresh-ListView $ListView[0].SelectedItems[0].Name}
		}

	$About_Click = {
		$AboutText = -join(
		$App.Vers,"{1}",
		'Created by Randy Turner - mailto:turner.randy21@yahoo.com',"{2}",
		'PS Movie Launcher was designed as a specialized task launcher.',"{2}",
		'Script Name: MovieLauncher.ps1',"{2}",   
		'Synopsis:',"{2}",
		'This script launches video files from a defineable location by use of a WinForm',"{1}",
		"it also includes the ability to download a selected video to the current user's","{1}",
		"'My Videos' folder. a GUI progress window will be displayed during file transfers.","{1}",
		"{0}","{1}",
		"Supported Video File Types: {3}","{1}",
		"{0}","{2}",
		'For additional help run the Powershell Get-Help cmdlet.' -f $("=" * 66),"`r`n",("`r`n"*2),(Convert-ArrayToString $VideoFileTypes))
		Show-AboutForm -AppName $App.Name -AboutText $AboutText -URL
	}

	$Back_Click = {
		if($script:previousItem)
			{Refresh-ListView $script:previousItem}
		else
			{Show-MessageBox -M 'Nowhere to go!' -T $Form1.Text -B OK -I Information}
		}

	$Home_Click = {
		if($CurrPath -ne "" -and $CurrPath -ne $Directory)
			{Refresh-ListView $Directory}
		else
			{Show-MessageBox -M 'You are currently Home' -T $Form1.Text -B OK -I Information}
		}

	$Host_Click = {Show-HostInfo}

	$Refresh_Click = {Refresh-ListView $Script:CurrPath}

	$Exit_Click = {$Form1.Close()}

	$Download_Click = {
		if($ListView[0].SelectedItems[0].Tag -eq $IconCatalogItems[1].Tag)
			{Invoke-CopyFile $ListView[0].SelectedItems[0].Name $($DownLoadTarget) -AsJob}
		else
			{Show-MessageBox -M 'Downloading an entire Directory is Not Supported!' -T $Form1.Text -B OK -I Warning}
		}

	$FontSettings_Click = {Invoke-FontDialog -Control $ListView[0] -FontMustExist -AllowSimulations -AllowVectorFonts}

	$OpenVideoFolder_Click = {Invoke-Item -Path $DownLoadTarget}

	$Find_Click = {
		$RV = Show-InputBox -Prompt 'Find?' -Title 'Search'
		if($RV -ne "") {$FindObj.Initialized = $False;Find-ListViewItem -SVI ([Ref]$FindObj) -Lvw $ListView[0] -Val $RV}
		}

	$FindNext_Click = {
		if($Script:FindObj.Initialized -eq $True)
			{Find-ListViewItem -SVI ([Ref]$FindObj) -Lvw $ListView[0]}
		else
			{Invoke-Command -ScriptBlock $Find_Click}
		}

	$Properties_Click = {
		if($ListView[0].SelectedItems[0].Count -gt 0)
			{
			Show-PropertySheet -TargetFile $ListView[0].SelectedItems[0].Name`
							   -IsDir ($ListView[0].SelectedItems[0].Tag -eq $IconCatalogItems[0].Tag)
			}
		else
			{Show-MessageBox -M 'No Item Selected!' -T $Form1.Text -B OK -Icon Warning}
		}

	$ToggleView_Click = {$ListView[0].View = Get-ActiveView}

	$Form_Load_StateCorrection = {
		$Form1.WindowState = $InitialFormWindowState
		Refresh-ListView $Directory
		}
	#endregion

	#region Common Control Variables
	$LvwCtxMenuItemsSize = New-Object -TypeName Drawing.Size -ArgumentList 219,32
	$ChkBoxLeftMost = 380
	$ChkBoxTop = 460
	#endregion

	#region Form Code
	$Form1.Text = $App.Name
	$Form1.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
	$Form1.Icon =[BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],'App Icon'),16)
	$Form1.ClientSize = New-Object -TypeName Drawing.Size -ArgumentList 692,510
	$Form1.MinimumSize = New-Object -TypeName Drawing.Size -ArgumentList 600,450
	$Form1.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen

	#region Populate Imagelists
	$DV = 0 
	for ($C = 0; $C -le 2; $C++) {$ImageList[$C].Images.Clear()}
	for ($C = 0; $C -le 1; $C++){
		$ImageList[0].Images.Add($IconCatalogItems[$C].Tag,(Get-Icon -Index $C -Size $IconSize))
		$ImageList[2].Images.Add($IconCatalogItems[$C].Tag,(Get-Icon -Index $C -Size 24))
		$DV++
		}
	for ($C = 0; $C -le $LvwCtxMenuitems.GetUpperBound(0); $C++){
		$ImageList[1].Images.Add($IconCatalogItems[$C + $DV].Tag,(Get-Icon -Index ($C+$DV) -Size 32))
		}
	#endregion 

	#region MainMenu 
	<#
	MainMenu is a Drop-Down menu designed to provide access to the
	various functions.
	#>
	$MainMenu.Visible = $True
	$MainMenu.Size = New-Object -TypeName Drawing.Size -ArgumentList 220,30
	$MainMenu.Items.AddRange($MainMenuItems)
	$MainMenuItemsSize = New-Object -TypeName Drawing.Size -ArgumentList 219,22

	$FileMenuText = @('Open\Play','Download File to My Videos','Find','Find Next','Properties','Exit')
	Set-MenuItem $FileMenuText $FileMenuItems $MainMenuItemsSize "FileMenuItem"
	$FileMenuItems[1].Enabled = `
	$FileMenuItems[1].Visible = !($SupressFileDownload)
	$FileMenuItems[0].Add_Click($LvwDouble_Click)
	$FileMenuItems[1].Add_Click($Download_Click)
	$FileMenuItems[2].Add_Click($Find_Click)
	$FileMenuItems[3].Add_Click($FindNext_Click)
	$FileMenuItems[4].Add_Click($Properties_Click)
	$FileMenuItems[5].Add_Click($Exit_Click)
	$FileMenuBar = New-Object -TypeName Windows.Forms.ToolStripSeparator

	$NaviMenuText = @('Back','Home','Refresh')
	Set-MenuItem $NaviMenuText $NaviMenuItems $MainMenuItemsSize 'NaviMenuItem'
	$NaviMenuItems[0].Add_Click($Back_Click)
	$NaviMenuItems[1].Add_Click($Home_Click)
	$NaviMenuItems[2].Add_Click($Refresh_Click)

	$ToolMenuText = @('Font Settings','Open Video Folder','Toggle View')
	Set-MenuItem $ToolMenuText $ToolMenuItems $MainMenuItemsSize 'ToolMenuItem'
	$ToolMenuItems[0].Add_Click($FontSettings_Click)
	$ToolMenuItems[1].Add_Click($OpenVideoFolder_Click)
	$ToolMenuItems[2].Add_Click($ToggleView_Click)
	$ToolMenuItems[1].Enabled = `
	$ToolMenuItems[1].Visible = !($SupressFileDownload)

	$HelpMenuText = @('About','Host Information')
	Set-MenuItem $HelpMenuText $HelpMenuItems $MainMenuItemsSize 'HelpMenuItem' -NoHotKeys -SetSizeOff
	$HelpMenuItems[0].Add_Click($About_Click)
	$HelpMenuItems[1].Add_Click($Host_Click)

	$MainMenuText = @('File','Navigation','Tools','Help')
	Set-MenuItem $MainMenuText $MainMenuItems $MainMenuItemsSize 'MainMenuItem' -NoHotKeys -SetSizeOff
	$MainMenuItems[0].DropDownItems.AddRange($FileMenuItems)
	$MainMenuItems[1].DropDownItems.AddRange($NaviMenuItems)
	$MainMenuItems[2].DropDownItems.AddRange($ToolMenuItems)
	$MainMenuItems[3].DropDownItems.AddRange($HelpMenuItems)
	$MainMenuItems[0].DropDownItems.Insert(4,$FileMenuBar)
	$Form1.Controls.Add($MainMenu)
	#endregion MainMenu

	#region Labels
	$Label1.Location = New-Object -TypeName Drawing.Point -ArgumentList 13,($ChkBoxTop + 29)
	$Label1.Text = ''
	$Label1.Width = 664
	$Label1.AutoEllipsis = $True
	$Label1.Anchor = Get-Anchor -B -L -R
	$Form1.Controls.Add($Label1)
	#endregion 

	#region LvwCtxMenuStrip
	$LvwCtxMenuStrip.Size = New-Object -TypeName Drawing.Size -ArgumentList 220,70
	$LvwCtxMenuStrip.Items.AddRange($LvwCtxMenuItems)
	#endregion

	#region LvwCtxMenuItems
	for($C = 0; $C -le $LvwCtxMenuItems.GetUpperBound(0); $C++){
		$LvwCtxMenuItems[$C].Text = $IconCatalogItems[$C + $DV].Tag
		$LvwCtxMenuItems[$C].Size = $LvwCtxMenuItemsSize
		$LvwCtxMenuItems[$C].Image = $ImageList[1].Images[$C]
		$LvwCtxMenuItems[$C].ImageAlign = 'MiddleLeft'
		$LvwCtxMenuItems[$C].ShowShortcutKeys = $True
		$LvwCtxMenuItems[$C].ShortcutKeys = Get-ShortcutKey -Mode $LvwCtxMenuItems[$C].Text
		}
	$LvwCtxMenuItems[0].Add_Click($LvwDouble_Click)
	$LvwCtxMenuItems[1].Add_Click($Download_Click)
	$LvwCtxMenuItems[2].Add_Click($Back_Click)
	$LvwCtxMenuItems[3].Add_Click($Home_Click)
	$LvwCtxMenuItems[4].Add_Click($Refresh_Click)
	$LvwCtxMenuItems[5].Add_Click($Find_Click)
	$LvwCtxMenuItems[6].Add_Click($FindNext_Click)
	$LvwCtxMenuItems[7].Add_Click($Properties_Click)
	$LvwCtxMenuItems[8].Add_Click($FontSettings_Click)
	$LvwCtxMenuItems[9].Add_Click($OpenVideoFolder_Click)
	$LvwCtxMenuItems[10].Add_Click($Exit_Click)
	$LvwCtxMenuItems[1].Visible = `
	$LvwCtxMenuItems[9].Visible = `
	$LvwCtxMenuItems[9].Enabled = !($SupressFileDownload)
	#endregion 

	#region Sepetator(s)
	$LvwCtxMenuStrip.Items.Insert(10, $LvwCtxMenuBars[1])
	$LvwCtxMenuStrip.Items.Insert(07, $LvwCtxMenuBars[0])
	#endregion

	#region listview[0] 
	$ListView[0].TabIndex = 0
	$ListView[0].View = Get-ActiveView
	$ListView[0].MultiSelect = $False
	$ListView[0].Size = New-Object -TypeName Drawing.Size -ArgumentList 668,458
	$ListView[0].Location = New-Object -TypeName Drawing.Point -ArgumentList 13,28
	$ListView[0].TileSize = New-Object -TypeName Drawing.Size -ArgumentList $IconSize,$IconSize
	$ListView[0].Font = New-Object -TypeName Drawing.Font -ArgumentList $FontName,$FontSize,(Get-SelectedFontStyle -FS $FontStyle)
	$ListView[0].SmallImageList = $imageList[2]
	$ListView[0].LargeImageList = $imageList[0]
	$ListView[0].ContextMenuStrip = $LvwCtxMenuStrip
	$ListView[0].UseCompatibleStateImageBehavior = $False
	$ListView[0].DataBindings.DefaultDataSourceUpdateMode = 0
	$ListView[0].Anchor = Get-Anchor -T -L -B -R
	$ListView[0].Add_DoubleClick($LvwDouble_Click)
	$Form1.Controls.Add($ListView[0])
	#endregion listvie[0] 
	#endregion Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $Form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$Form1.Add_Load($Form_Load_StateCorrection)
	#Show the Form
	[Void]$Form1.ShowDialog()
}

#Call the Main function
if($AbEnd -ne $True) {Show-MainForm}