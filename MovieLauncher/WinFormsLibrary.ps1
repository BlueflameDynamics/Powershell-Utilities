<#
.NOTES
Name:		WinFormsLibrary.ps1
Author:		Randy Turner
Version:	1.4
Date:		08/06/2022
Revision History:
v1.4 - 08/06/2022 - Modified Show-MessageboxEx to fix DisplayDefaultDesktop issue
v1.3 - 03/14/2022 - Added New Alias
v1.2 - 03/14/2021 - Added Form & Utility Regions & Screen functions.
v1.1 - 03/15/2020 - Added New-ObjectArray function
v1.0 - 06/05/2014 - Original Release

.SYNOPSIS
Provides a wrapper for utility functions used with WinForms.
#>

#region Common Variables
Add-Type -A System.Drawing
Import-Module -Name .\Exists.ps1 -Force
$ImageRes = $Env:SystemRoot+'\System32\imageres.dll'
$PSCoreDefFont = [System.Drawing.Font]::New('Segoe UI',9,[System.Drawing.FontStyle]::Regular)
#endregion

#region Add BlueflameDynamics.IconTools Class
$DLLPath = '.\BlueflameDynamics.IconTools.dll'
if((Test-Exists -Mode File -Location $DLLPath) -eq $True){Add-Type -Path $DLLPath}
#endregion

#region Forms
<#
.NOTES
Name:		Invoke-FontDialog Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Invoke a WinForm FontDialog

.PARAMETER Control
Required, Windows Control to interact with the FontDialog

.PARAMETER MinSize
Optional, Minimum font size in points, Defaults to 9

.PARAMETER MaxSize
Optional, Maximum font size in points, Defaults to 24

.PARAMETER Showcolor
Optional, Switch if present allows Font color

.PARAMETER FontMustExist
Optional, Switch if present the Font must exist on the host

.PARAMETER FixedPitchOnly
Optional, Switch if present only Fixed Pitch Fonts will be listed

.PARAMETER AllowSimulations
Optional, Switch if present allows graphics device interface (GDI) font simulations

.PARAMETER AllowVerticalFonts
Optional, Switch if present allows listing Vertical Fonts 

.PARAMETER AllowVectorFonts
Optional, Switch if present allows listing Vector Fonts

.PARAMETER ShowEffects
Optional, Switch if present allows the user to specify strikethrough, underline, and text color options

.EXAMPLE
$MainForm.Font = Invoke-FontDialog -Control $MainForm -FixedPitchOnly
This example returns a Selected Fixed Pitch font or leaves the font unchanged if the dialog is cancelled

.EXAMPLE
$ListView.Font = Invoke-FontDialog -Control $ListView -FixedPitchOnly -ShowEffects -MinSize 6
This example returns a Selected Fixed Pitch font, allow effects, & set the minimum font size to 6pts 
#>
function Invoke-FontDialog{
	param(
		[Parameter(Mandatory)][Windows.Forms.Control]$Control,
		[Parameter()][Int]$MinSize = 9,
		[Parameter()][Int]$MaxSize = 24,
		[Parameter()][Switch]$Showcolor,
		[Parameter()][Switch]$FontMustExist,
		[Parameter()][Switch]$FixedPitchOnly,
		[Parameter()][Switch]$AllowSimulations,
		[Parameter()][Switch]$AllowVerticalFonts,
		[Parameter()][Switch]$AllowVectorFonts,
		[Parameter()][Switch]$ShowEffects)

	$FontDialog = [Windows.Forms.FontDialog]::New()
	$FontDialog.Showcolor = $Showcolor
	$FontDialog.FontMustExist = $FontMustExist
	$FontDialog.FixedPitchOnly = $FixedPitchOnly
	$FontDialog.AllowSimulations = $AllowSimulations
	$FontDialog.AllowVerticalFonts = $AllowVerticalFonts
	$FontDialog.AllowVectorFonts = $AllowVectorFonts
	$FontDialog.ShowEffects = $ShowEffects
	$FontDialog.minSize = $MinSize
	$FontDialog.maxSize = $MaxSize
	$FontDialog.Font = $Control.Font
	$FontDialog.Font = $Control.Font
	$RV = $FontDialog.ShowDialog()
	if($RV -eq [System.Windows.Forms.DialogResult]::OK){
		$Control.Font = $FontDialog.Font
		$Control.Refresh
	}
	return $RV
}

<#
.NOTES
Name:		Show-AboutForm Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Display an AboutBox

.PARAMETER AppName
Required, Application Name to be used within the AboutBox

.PARAMETER AboutText
Required, Text to be used within the AboutBox's RichTextBox

.PARAMETER FormWidth
Optional, Width of AboutBox Form, Default is 570

.PARAMETER FormHeight
Optional, Height of AboutBox Form, Default is 275

.PARAMETER DetectUrls
Optional, Switch if present causes the URLs within the About Text to be detected, marked, & Clickable

.EXAMPLE
$About_Click ={
$AboutText = ("<Some Help Test>") 
Show-AboutForm -AppName "AboutBox Test" -DetectUrls -AboutText $AboutText
}
This example displays the default About
#>
function Show-AboutForm{
	param (
		[Parameter(Mandatory)][Alias('N')][String]$AppName,
		[Parameter(Mandatory)][Alias('T')][String]$AboutText,
		[Parameter()][Alias('W')][Int]$FormWidth = 570,
		[Parameter()][Alias('H')][Int]$FormHeight = 275,
		[Parameter()][Alias('URL')][Switch]$DetectUrls)

	Add-Type -A System.Windows.Forms
	#Add objects for About
	$FormAbout = [Windows.Forms.Form]::New()
	$RtbAbout = [Windows.Forms.RichTextBox]::New()
	$InitialFormWindowState = [Windows.Forms.FormWindowState]::Normal
	
	#About Form
	$FormAbout.Name = 'FormAbout'
	$FormAbout.AutoScroll = $True
	$FormAbout.ClientSize = [Drawing.Size]::New($FormWidth,$FormHeight)
	$FormAbout.DataBindings.DefaultDataSourceUpdateMode = 0
	$FormAbout.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
	$FormAbout.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$FormAbout.Text = ' About ' + $AppName
	$FormAbout.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($ImageRes,76,24)
	$FormAbout.MaximizeBox = `
	$FormAbout.MinimizeBox = $False
	$FormAbout.ShowInTaskbar = $False

	$RtbAbout.Name = 'RtbAbout'
	$RtbAbout.Size = [Drawing.Size]::New($FormWidth-13,$FormHeight-13)
	$RtbAbout.Font = $PSCoreDefFont
	$RtbAbout.Location = [Drawing.Point]::New(13,13)
	$RtbAbout.Anchor = Get-Anchor -T -L -B -R
	$RtbAbout.BackColor = [Drawing.SystemColors]::Window
	$RtbAbout.BorderStyle = 0
	$RtbAbout.DataBindings.DefaultDataSourceUpdateMode = 0
	$RtbAbout.DetectUrls = $DetectUrls
	$RtbAbout.ReadOnly = $True
	$RtbAbout.Cursor = Get-Cursor -Mode Default
	$RtbAbout.TabIndex = 0
	$RtbAbout.TabStop = $False
	$RtbAbout.Text = -Join $FormAbout.Text, $AboutText
	
	#Handles clicking the links in about form
	$RtbAbout.add_LinkClicked({ Start-Process -FilePath $_.LinkText })
	$FormAbout.Controls.Add($RtbAbout)
	[Void]$FormAbout.ShowDialog()
}

<#
.NOTES
Name:		Show-InputBox Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to display an InputBox & Return the entered text
A $null is returned upon cancellation

.PARAMETER Prompt (Alias <P>)
Required, Text to Prompt for Input

.PARAMETER Title (Alias <T>)
Optional, Text of InputBox Title

.PARAMETER Default (Alias <DV>)
Optional, Text of InputBox Default Value

.EXAMPLE
$RV = Show-InputBox -Prompt "Find?" -Title "Search";if ($RV -ne "") { Find-ListViewItem $ListView[0] $RV }
This example prompts for input of a search value
#>
function Show-InputBox{
	param (
		[Parameter(Mandatory)][Alias('P')][String]$Prompt,
		[Parameter()][Alias('T')][String]$Title = '',
		[Parameter()][Alias('DV')][String]$Default = '')

	If ($Title.Length -eq 0) { $Title = ' ' }

	return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}

<#
.NOTES
Name:		Show-MessageBoxEx Function
Author:		Randy Turner
Version:	2.0
Date:		08/06/2022
Revision History:
	V2.0 - 08/06/2022 - Removed Options parameter, Added hidden form
	V1.0 - 12/05/2021 - Inital release

.SYNOPSIS
Provides a wrapper for function used to Display a MessageBox & get the button selected.
Includes parameters to use advanced MessageBox functionality.

.PARAMETER Message - Alias (M)
Required, Message Text

.PARAMETER Title - Alias (T)
Optional, Form Title Text

.PARAMETER Buttons - Alias (B)
Optional, a value for the Message Box Buttons to include

.PARAMETER Icon - Alias (I)
Optional, Message Box Icon

.PARAMETER DefaultButton - Alias (D)
Optional, a value for the Message Box Default Button

.EXAMPLE
$RV = Show-MessageBoxEx -M "This is a Test!" -T "Test Title" -B YesNoCancel -I Warning -D Button1
This example displays a MessageBox returns the selected button
#>
function Show-MessageBoxEx{
	param(
	[Parameter()][Alias('O')][System.Windows.Forms.Form]$Owner = $Null,
	[Parameter(Mandatory)][Alias('M')][String]$Messsage,
	[Parameter()][Alias('T')][String]$Title = '',
	[Parameter()][Alias('B')]
		[System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
	[Parameter()][Alias('I')]
		[System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::None,
	[Parameter()][Alias('D')]
		[System.Windows.Forms.MessageBoxDefaultButton]$DefaultButton = [System.Windows.Forms.MessageBoxDefaultButton]::Button1)

	#Loads the WinForm Assembly$HF  = New-Object -TypeName Windows.Forms.Form
	Add-Type -A System.Windows.Forms
	
	#Display the message with input
	$Flag = $False
	if($Owner.Length -eq 0){
		#Create a Hidden Form to simulate the Option parameter: DefaultDesktopOnly,
		#which doesn't always work reliably in Powershell. Forces Messagebox as TopMost Form.
		#region Hidden Form
		$HF = [Windows.Forms.Form]::New()
		$HF.Opacity = 0
		$HF.Visible = `
		$HF.ShowInTaskbar = $False
		$HF.Enabled = `
		$HF.TopMost = $True
		$HF.Size = [Drawing.Size]::New(50,50)
		$HF.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
		$HF.Add_Shown({$This.Hide()})
		$Owner = $HF
		$Flag = $True
		[Void]$HF.Show()}
		#endregion
	$DR = [System.Windows.Forms.MessageBox]::Show($Owner,$Messsage,$Title,$Buttons,$Icon,$DefaultButton)
	If($Flag){$HF.Close()}
	$DR
}

#Assign an Alias
Set-Alias -Name Show-MessageBox -Value Show-MessageBoxEx

<#
.NOTES
Name:		Show-HelpForm Function
Author:		Randy Turner
Version:	1.3
Date:		06/01/2022
Revision History:
	V1.0 - 06/05/2014 - Original Release
	V1.1 - 12/04/2019 - Enhanched Text-to-Speech
	V1.2 - 12/11/2019 - Converted Text-to-Speech to Async & added an escape
	V1.3 - 06/01/2022 - Added gender support to Text-to-Speech

.SYNOPSIS
Provides a wrapper for fumction used to Display a Help Window

.PARAMETER AppName
Required, Application Name to be used within the AboutBox

.PARAMETER HelpText
Required, Text to be used within the HelpForm's RichTextBox

.PARAMETER FormWidth
Optional, Width of Help Form, Default is 510

.PARAMETER FormHeight
Optional, Height of Help Form, Default is 560

.PARAMETER DetectUrls - Alias: URL
Optional, Switch if present causes the URLs within the About Text to be detected, marked, & Clickable

.PARAMETER ReadHelpFile - Alias: File
Optional, Switch if present causes the HelpText parameter to be treated as a text filename
The file is imported and displayed in the Help Form.

.EXAMPLE
$Help_Click ={
$HelpText = ("<Some Help Test>") 
Show-HelpForm -AppName "Help Test" -DetectUrls -HelpText $HelpText
}
This example displays the default About
#>
function Show-HelpForm{

	param (
			[Parameter(Mandatory)][String]$AppName,
			[Parameter(Mandatory)][String]$HelpText,
			[Parameter()][Int]$FormWidth = 510,
			[Parameter()][Int]$FormHeight = 560,
			[Parameter()][Alias('URL')][Switch]$DetectUrls,
			[Parameter()][Alias('File')][Switch]$ReadHelpFile,
			[Parameter()][Alias('Read')][Switch]$ReadText)

	Add-Type -A System.Windows.Forms

	#region Utility Functions
	function Get-ShortcutKey{
	param(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Font Settings','Toggle Voice')]
		[String]$Mode)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$ValidModes = $MyParam['Mode'].Attributes.ValidValues
	$WFK = [Windows.Forms.Keys]

	switch([Array]::IndexOf($ValidModes,$Mode)){
		 0 {$WFK::Control -bor $WFK::F}
		 1 {$WFK::Control -bor $WFK::V}
		 Default {$WFK::None}
		}}

	Function Set-MenuItem{
	Param(
		[Parameter(Mandatory)][String[]]$Labels,
		[Parameter(Mandatory)][Windows.Forms.ToolStripMenuItem[]]$MenuItems,
		[Parameter(Mandatory)][Drawing.Size]$ItemSize,
		[Parameter(Mandatory)][String]$ItemPrefix,
		[Parameter()][String[]]$HotKeys = @(),
		[Parameter()][Switch]$SetSizeOff,
		[Parameter()][Switch]$NoHotKeys)
	
	for($C=0;$C -le $MenuItems.GetUpperBound(0);$C++){
		$MenuItems[$C].Name = $ItemPrefix + ($C+1)
		$MenuItems[$C].Text = $Labels[$C]
		if(!$SetSizeOff){$MenuItems[$C].Size = $ItemSize}
		if(!$NoHotKeys){
			$MenuItems[$C].ShortcutKeys =`
			Get-ShortcutKey -Mode $(if($Hotkeys.Length -eq 0) {$Labels[$C]}else{$HotKeys[$C]})
			}
		}
	}
	#endregion
	
	#region Add objects for Help
	$FrmHelp = [Windows.Forms.Form]::New()
	$RtbHelp = [Windows.Forms.RichTextBox]::New()
	$InitialFormWindowState = [System.Windows.Forms.FormWindowState]::Normal
	$Help_MainMenu = [Windows.Forms.MenuStrip]::New()
	$Help_MainMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 1
	$Help_ToolMenuItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count 2
	#endregion 

	#region Help form}
	if($Host.Version.Major -gt 5){$FormWidth += 75} 
	$FrmHelp.AutoScroll = $True 
	$FrmHelp.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Size = [Drawing.Size]::New($FormWidth,$FormHeight)
	$FrmHelp.MinimumSize = $System_Drawing_Size
	$FrmHelp.AutoScalemode = [System.Windows.Forms.AutoScaleMode]::Dpi
	$FrmHelp.AutoSize = $True
	$FrmHelp.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowOnly
	$FrmHelp.Name = 'FrmHelp'
	$FrmHelp.ShowInTaskbar = $False
	$FrmHelp.Text = -join $AppName,' - Help'
	$FrmHelp.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
	$FrmHelp.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$FrmHelp.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($ImageRes, 94, 16)
	$FrmHelp.MaximizeBox = $False
	$FrmHelp.Controls.Add($Help_MainMenu)
	#endregion

	#region Main Menu
	$Help_MainMenu.Name = 'Help_MainMenu'
	$Help_MainMenu.Visible = $True
	$Help_MainMenu.Size = [Drawing.Size]::New(220,30)
	$Help_MainMenu.Items.AddRange($Help_MainMenuItems)
	$Help_MainMenuItemsSize = [Drawing.Size]::New(219,22)
	$Help_MainMenuText = @('Tools')
	Set-MenuItem $Help_MainMenuText $Help_MainMenuItems $Help_MainMenuItemsSize 'Help_MainMenuItem' -NoHotKeys -SetSizeOff
	
	$Help_ToolMenuText = @('Font Settings','Toggle Voice')
	Set-MenuItem $Help_ToolMenuText $Help_ToolMenuItems $Help_MainMenuItemsSize 'Help_ToolMenuItem' 
	$Help_MainMenuItems[0].DropDownItems.AddRange($Help_ToolMenuItems)
	$Help_FontSettings_Click = {Invoke-FontDialog -Control $RtbHelp -FontMustExist -AllowSimulations -AllowVectorFonts}
	$Help_ToggleVoice_Click = {$Help_ToolMenuItems[1].Checked = !$Help_ToolMenuItems[1].Checked}
	$Help_ToolMenuItems[0].Add_Click($Help_FontSettings_Click)
	$Help_ToolMenuItems[1].Add_Click($Help_ToggleVoice_Click)
	#endregion

	#region Rich Textbox
	$RtbHelp.Size = [Drawing.Size]::New($FormWidth-45,$FormHeight*.78)
	$RtbHelp.Anchor = Get-Anchor -T -L -B -R 
	$RtbHelp.BackColor = [Drawing.SystemColors]::Window
	$RtbHelp.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$RtbHelp.Location = [Drawing.Point]::New(13,30)
	$RtbHelp.Name = 'RtbHelp'
	$RtbHelp.ReadOnly = $True
	$RtbHelp.SelectionProtected = $True
	$RtbHelp.Cursor = Get-Cursor Default
	$RtbHelp.TabIndex = 0
	$RtbHelp.TabStop = $False
	$RtbHelp.DetectUrls = $DetectUrls

	if($ReadHelpFile.IsPresent){
		if((Test-Exists -Mode File -Location $Helptext) -eq $False){
			Show-MessageBox `
				-M ('Help File: {0} Not Found!' -f $HelpText) `
				-T 'Help File Not Found!' -B Ok -I Warning
		return $Null
		}
	$Lines = Get-Content -Path $Helptext
	$HelpText = ''
	foreach($Line in $Lines){$HelpText+="{0}`r`n" -f $Line}
	}

	$RtbHelp.Text = -join $AppName,' - Help',$HelpText

	#Handles clicking of links in help document
	$RtbHelp.Add_LinkClicked({Invoke-Expression -Command "Start $($_.LinkText)"})
	$FrmHelp.Controls.Add($RtbHelp)
	#endregion

	#region Buttons
	$BtnWidth = 80
	$BTT = @('&Stop','&Read Text','E&xit')
	$Buttons = New-ObjectArray -TypeName System.Windows.Forms.Button -Count 3
	$BC = $Buttons.Count
	for($C=0;$C -lt $Buttons.Count;$C++){
		$Buttons[$C].Anchor = Get-Anchor -B -R
		$Buttons[$C].Name = 'Btn'+$BTT[$C]
		$Buttons[$C].Size = [System.Drawing.Size]::New($BtnWidth,30)
		$Buttons[$C].Left = $RtbHelp.Right - ($BtnWidth*$BC)
		$Buttons[$C].Location = [System.Drawing.Point]::New($Buttons[$C].Left,$FrmHelp.Height*.86)
		$Buttons[$C].Text = $BTT[$C]
		$Buttons[$C].Enabled = `
		$Buttons[$C].Visible = $True
		$Buttons[$C].UseVisualStyleBackColor = $True
		$FrmHelp.Controls.Add($Buttons[$C])
		$BC--
	}
	$Buttons[0].Enabled = `
	$Buttons[0].Visible = $False
	$Buttons[0].Location = $Buttons[1].Location
	if(!$ReadText.IsPresent){
		$Buttons[1].Enabled = `
		$Buttons[1].Visible = $False
	}
	$BtnRead_Click = {
		$RtbHelp.Cursor = Get-Cursor -Mode WaitCursor
		Add-Type -A System.Speech
		$SayIt = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
		$Voice = if($Help_ToolMenuItems[1].Checked){'Male'}else{'Female'}
		$SayIt.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::$Voice)
		$SayIt.SpeakAsync($RtbHelp.Text)
		$Buttons[1].Enabled = `
		$Buttons[2].Enabled = `
		$Buttons[1].Visible = $False
		$Buttons[0].Enabled = `
		$Buttons[0].Visible = $True
		$I = 0
		Do{<#Wait While Speaking Loop#>
			If($I % (1000*5) -eq 0){
				[System.Windows.Forms.Application]::DoEvents()
				$I = 0}
			$I++
		}
		While($SayIt.State -eq [System.Speech.Synthesis.SynthesizerState]::Speaking)
		$RtbHelp.Cursor = Get-Cursor -Mode Default
		$Buttons[1].Enabled = `
		$Buttons[2].Enabled = `
		$Buttons[1].Visible = $True
		$Buttons[0].Visible = $False
	}
	$BtnExit_Click = {$FrmHelp.Close()}
	$BtnStop_Click = {$SayIt.SpeakAsyncCancelAll()}
	$Buttons[0].Add_Click($BtnStop_Click)
	$Buttons[1].Add_Click($BtnRead_Click)
	$Buttons[2].Add_Click($BtnExit_Click)
	#endregion

	[Void]$FrmHelp.ShowDialog()
}

<#
.NOTES
Name:	Invoke-OpenFileDialog Function
Author:  Randy Turner
Version: 1.0
Date:	03/22/2017

.SYNOPSIS
Provides a wrapper fumction used to Display an OpenFileDialog Control 
and return either the contents of the selected file, the selected filename(s), 
or an empty string upon cancellation.

.PARAMETER ReturnMode. Alias: Mode
Optional, Used to specify the desired output: Contents(default)/Filename/Multiple

.PARAMETER Title  Alias: T
Optional, String used to set the OpenFileDialog.Title, defualt to $NULL.

.PARAMETER FileFilter  Alias: F
Optional, Filter String used to set the OpenFileDialog.Filter.
Defaults to "Text files (*.txt)|*.txt|All files (*.*)|*.*"

.PARAMETER FilterIndex  Alias: FI
One based index used to select the desired default file type

.PARAMETER InitialDirectory  Alias: Dir
used to set the Initial Directory of the control

.PARAMETER RestoreDirectory  Alias: RD
Determines wether the previously selected location is restored upon exit

.PARAMETER ShowHelp. Alias: ShowHelp
Determines if the Help button is shown, Overridden when
running under the 'Default Host' to True

.EXAMPLE
$FileContents = Invoke-OpenFileDialog
This example displays the OpenFileDialog and returns the selected file contents

.EXAMPLE
$Filename = Invoke-OpenFileDialog -ReturnMode Filename
This example displays the OpenFileDialog and returns the selected filename

.EXAMPLE
$Filename = Invoke-OpenFileDialog -ReturnMode Multiple
This example displays the OpenFileDialog and returns an array of selected filenames
#>
function Invoke-OpenFileDialog{
	param(
		[Parameter()][Alias('Mode')]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('Contents','Filename','Multiple')]
			[String]$ReturnMode = 'Contents',
		[Parameter()][Alias('T')][String]$Title=$null,
		[Parameter()][Alias('FN')][String]$File='',
		[Parameter()][Alias('F')][String]$FileFilter='Text files (*.txt)|*.txt|All files (*.*)|*.*',
		[Parameter()][Alias('FI')][Int]$FilterIndex=1,
		[Parameter()][Alias('Dir')][String]$InitialDirectory='.',
		[Parameter()][Alias('RD')][Switch]$RestoreDirectory,
		[Parameter()][Alias('SH')][Switch]$ShowHelp)

	Add-Type -AssemblyName System.Windows.Forms

	<# 
	If Running under the Default Host $ShowHelp 
	MUST BE True or the Control Hangs.
	#>
	if($Host.Name -eq 'Default Host'){$ShowHelp = $True}

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$ReturnModes = $MyParam['ReturnMode'].Attributes.ValidValues
	
	$RV = ''
	$ModeIdx = [Array]::IndexOf($ReturnModes,$ReturnMode)
	$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::New()
	$OpenFileDialog.InitialDirectory = $InitialDirectory
	$OpenFileDialog.RestoreDirectory = $RestoreDirectory
	$OpenFileDialog.AutoUpgradeEnabled = $True
	$OpenFileDialog.Filter = $FileFilter
	$OpenFileDialog.FilterIndex = $FilterIndex
	$OpenFileDialog.Title = $Title
	$OpenFileDialog.FileName = $File
	$OpenFileDialog.Multiselect = ($ModeIdx -eq 2)
	$OpenFileDialog.ShowHelp = $ShowHelp
	if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
		switch($ModeIdx){
			0 {$RV = Get-Content -Path $OpenFileDialog.FileName}
			1 {$RV = $OpenFileDialog.FileName}
			2 {$RV = $OpenFileDialog.FileNames}
		}
	}
	return $RV
}

<#
.NOTES
Name:		Show-HostInfo Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2022

.SYNOPSIS
This function displays a MessageBox of information about the current Host
#>
function Show-HostInfo{
	param(
		[Parameter()][Alias('O')][System.Windows.Forms.Form]$Owner=$Form1,
		[Parameter()][Alias('A')][String]$App=$App.Name)
	$Platform = if([Environment]::Is64BitProcess){64}else{32}
	$Mask ="Name : {0}`nVersion: {1}`nCurrentCulture: {2}`nCurrentUICulture: {3}`nArchitecture: {4}-bit"
	$Msg = $Mask -f $Host.Name,$Host.Version,$Host.CurrentCulture,$Host.CurrentUICulture,$Platform
	Show-Messagebox -Owner $Owner -M $Msg -T "$App - Host Information" -B OK -I Information
}
#endregion

#region Utility
<#
.NOTES
Name:		Get-Anchor Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to get a WinForm Anchor value.

.PARAMETER Top
Alias: T
Optional, causes the TOP Anchor to be included.

.PARAMETER Bottom
Alias: B
Optional, causes the BOTTOM Anchor to be included.

.PARAMETER Left
Alias: L
Optional, causes the LEFT Anchor to be included.

.PARAMETER Right
Alias: R
Optional, causes the RIGHT Anchor to be included.

.EXAMPLE
$Textbox1.Anchor = Get-Anchor -T -L -B -R
This example returns an Anchor value for all four Anchors.

.EXAMPLE
$Textbox1.Anchor = Get-Anchor -T -L
This example returns an Anchor value for the Top & Left Anchors.

.EXAMPLE
$Button1.Anchor = Get-Anchor
This example returns an Anchor value clearing all Anchors.
#>
function Get-Anchor{
	param(
		[Parameter()][Alias('T')][Switch]$Top,
		[Parameter()][Alias('B')][Switch]$Bottom,
		[Parameter()][Alias('L')][Switch]$Left,
		[Parameter()][Alias('R')][Switch]$Right)
	
	$Anchors = @(0,0,0,0)
	$Anchors[0] = if($Top.IsPresent){[Windows.Forms.AnchorStyles]::Top}
	$Anchors[1] = if($Bottom.IsPresent){[Windows.Forms.AnchorStyles]::Bottom} 
	$Anchors[2] = if($Left.IsPresent){[Windows.Forms.AnchorStyles]::Left}
	$Anchors[3] = if($Right.IsPresent){[Windows.Forms.AnchorStyles]::Right}	 
	return [Windows.Forms.AnchorStyles]`
		$($Anchors[0] -bor $Anchors[1] -bor $Anchors[2] -bor $Anchors[3])
}

<#
.NOTES
Name:		Get-Cursor Function
Author:		Randy Turner
Version:	1.0
Date:		06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to set a WinForm Cursor

.PARAMETER Mode
Required, Cursor type to return 

.EXAMPLE
$MainForm.Cursor = Get-Cursor -Mode AppStarting
This example returns an AppStarting Cursor

.EXAMPLE
$MainForm.Cursor = Get-Cursor -Mode WaitCursor
This example returns the WaitCursor

.EXAMPLE
$MainForm.Cursor = Get-Cursor Default
This example returns the Default Cursor
#>
function Get-Cursor{
	param(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet(
			'AppStarting','Arrow','Cross','Default','Hand','Help','HSplit','IBeam','No','NoMove2D',
			'NoMoveHoriz','NoMoveVert','PanEast','PanNE','PanNorth','PanNW','PanSE','PanWest',
			'PanSouth','PanSW','SizeAll','SizeNESW','SizeWE','UpArrow','VSplit','WaitCursor')]
		[String]$Mode)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$Modes = $MyParam['Mode'].Attributes.ValidValues 

	$WFC = [Windows.Forms.Cursors]
	
	switch([Array]::IndexOf($Modes, $Mode)){ 
		#Set Cursor
		00 {$WFC::AppStarting}
		01 {$WFC::Arrow}
		02 {$WFC::Cross}
		04 {$WFC::Hand}
		05 {$WFC::Help}
		06 {$WFC::HSplit}
		07 {$WFC::IBeam}
		08 {$WFC::No}
		09 {$WFC::NoMove2D}
		10 {$WFC::NoMoveHoriz}
		11 {$WFC::NoMoveVert}
		12 {$WFC::PanEast}
		13 {$WFC::PanNE}
		14 {$WFC::PanNorth}
		15 {$WFC::PanNW}
		16 {$WFC::PanSE}
		17 {$WFC::PanSouth}
		18 {$WFC::PanSW}
		19 {$WFC::PanWest}
		20 {$WFC::SizeAll}
		21 {$WFC::SizeNESW}
		22 {$WFC::SizeWE}
		23 {$WFC::UpArrow}
		24 {$WFC::VSplit}
		25 {$WFC::WaitCursor}
		Default {$WFC::Default} #03
	}
}

<#
.NOTES
Name:		New-ObjectArray Function
Author:		Randy Turner
Version:	1.0
Date:		07/15/2020

.SYNOPSIS
Provides a wrapper for fumction used to create an array.of objects

.DESCRIPTION
This function calls the New-Object cmdlet recursively to build an array of objects 
The -TypeName parameter corresponsponds to the New-Object TypeName and the -Value 
parameter indicates the number of instances of the requested object type to include
in the returned array.

.PARAMETER TypeName - Alias: T
Required, Type of Object to include in array.

.PARAMETER Count - Alias: C
Required, Number of object instances to be included.

.EXAMPLE
$MainMenuItems = New-ControlArray -TypeName Windows.Forms.ToolStripMenuItem -Value 7
This example returns an Array of Windows.Forms.ToolStripMenuItem Items.

.EXAMPLE
$Buttons = New-ControlArray -T Windows.Forms.Button -V 3
This example returns an Array of Windows.Forms.Button Items.
#>
function New-ObjectArray{
	param(
		[Parameter(Mandatory)][Alias('T')][String]$TypeName,
		[Parameter(Mandatory)][Alias('C')][Int]$Count) 
	@(for ($C = 1;$C -le $Count;$C++) {New-Object -TypeName $TypeName})
}

#Assign an Alias
Set-Alias -Name New-ControlArray -Value New-ObjectArray

<#
.NOTES
Name:		RepositionTo-CenterScreen Function
Author:		Randy Turner
Version:	1.0
Date:		03/13/2021

.SYNOPSIS
Provides a function used to reposition a WinForm to Center Screen. 

.PARAMETER Form
Required, Form to relocate.

.EXAMPLE
RepositionTo-CenterScreen -Form $MainForm
This example relocate the $MainForm
#>
function RepositionTo-CenterScreen(){
param([Parameter(Mandatory)][Alias('F')][Windows.Forms.Form]$Form)
		$Screen =  [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
		$Form.Top = ($Screen.Height / 2) - ($Form.Height / 2)
		$Form.Left = ($Screen.Width / 2) - ($Form.Width / 2)
}

<#
.NOTES
Name:		ScaleToScreen Function
Author:		Randy Turner
Version:	1.0
Date:		03/13/2021

.SYNOPSIS
Provides a fumction used to rescale a WinForm to the host Screen.
Requires Boolean $ScaleToScreen = $True to Enable,Assumes $False if not declared.

.PARAMETER Form
Required, Form to rescale.

.EXAMPLE
RepositionTo-CenterScreen -Form $MainForm
This example relocate the $MainForm
#>
#region Default Screen Size
$DefaultScreen = [System.Drawing.Size]::New(1920,1040)
#endregion

function ScaleToScreen(){
param(
	[Parameter(Mandatory)][Alias('F')][Windows.Forms.Form]$Form,
	[Parameter()][Alias('C')][Switch]$CenterScreen)

	[System.Drawing.SizeF] $BaseScaleFactor = [System.Drawing.SizeF]::New(1,1) 
	# Scale our form to look like it did when we designed it($DefaultScreen).
	# This adjusts between the screen resolution of the design computer and the active computer.
	$CurScreenWidth  = [System.Windows.Forms.Screen]::FromControl($Form).WorkingArea.Width
	$CurScreenHeight = [System.Windows.Forms.Screen]::FromControl($Form).WorkingArea.Height
	[float] $ScaleFactorWidth = [float]$CurScreenWidth / $DefaultScreen.Width
	[float] $ScaleFactorHeigth = [float]$CurScreenHeight / $DefaultScreen.Height
	[System.Drawing.SizeF] $ScaleFactor = [System.Drawing.SizeF]::New($ScaleFactorWidth,$ScaleFactorHeigth)
	If($ScaleFactor -ne $BaseScaleFactor){'ScaleFactor: {0}' -f $ScaleFactor|Out-Host}
	$Form.Scale($ScaleFactor)

	# If you want to center the resized screen.
	if($CenterScreen.IsPresent){RepositionTo-CenterScreen -Form $Form}
	return $ScaleFactor
}
#endregion