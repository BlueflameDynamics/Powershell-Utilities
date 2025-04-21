Import-Module -Name .\Exists.ps1 -force
Import-Module -Name .\BlueflameDynamics.IconTools.dll -force
Import-Module -Name .\WinFormsLibrary.ps1 -force

#region Local Enums
Enum ColorPickerSortMode{
	ArgbHexValueAscending
	ArgbHexValueDescending
	BlueChannelAscending
	BlueChannelDescending
	BrightnessAscending
	BrightnessDescending
	GreenChannelAscending
	GreenChannelDescending
	HueAscending
	HueDescending
	NameAscending
	NameDescending
	RedChannelAscending
	RedChannelDescending
	SaturationAscending
	SaturationDescending
}

Enum FormLabel{
	OldColorLabel
	OldColorPatch
	NewColorLabel
	NewColorPatch
}
#endregion

#region Get Expanded Web ColorInfo objects with Properties for Sort.
# Get-ColorInfo returns an ArrayList with custom methods.
$ArlColors = [System.Collections.ArrayList]::New()
$ArlColors = .\Get-ColorInfo.ps1 -ColorSet Web -OutputType ColorInfoTable
$ArlColors.RemoveAt(0) #Remove Tranparent
#endregion

<#
.NOTES
Name:	Show-ColorPicker Function
Author:  Randy Turner
Version: 1.0f
Date:	12/20/2022
Revision History -----------------------------------------------------------------------
v1.0  - 06/20/2019 - Original release
v1.0a - 06/21/2019 - Added additional comments for clarity & a minor bug fix
v1.0b - 06/24/2019 - Added Tab support for color grid buttons
v1.0c - 07/03/2019 - Added support for sort mode listbox as sorted list
v1.0d - 08/30/2019 - removed positional parameter references
v1.0e - 11/25/2019 - simplified code
v1.0f - 12/20/2022 - added option to return ColorInfo object.
----------------------------------------------------------------------------------------

.SYNOPSIS
Provides a wrapper for function used to Display a WinForms ColorPicker.

.DESCRIPTION
This script provides a WinForms ColorPicker Dialog that will display the 141 colors of the
Windows Web Color Set. Two larger color swatches showing the Old and New color selections
are included. There are 16 color sort options provided for Name, HSL values, & ARGB channels.
upon selection a custom PSObject is returned with the DialogResult and the selected color.
	----------------------------------------------------------------------------------------
	Security Note: This is an unsigned script, Powershell security may require you run the
	Unblock-File cmdlet with the Fully qualified filename before you can run this script,
	assuming PowerShell security is set to RemoteSigned.
	---------------------------------------------------------------------------------------- 

.PARAMETER SortMode Alias: S
Optional, a custom [ColorPickerSortMode] object that sets the Color Sort Mode.

.PARAMETER CurrentColor Alias: C
Optional, a [System.Drawing.Color] that represents the current color to be changed

.PARAMETER HideTransparent Alias: HT
Optional, switch that will cause the color "Transparent" to be disabled & hidden.

.PARAMETER ReturnColorInfo Alias: RC
Optional, switch that will cause the color returned to be expressed as a PsCustomObject of [BlueflameDynamics.ColorInfo]

.EXAMPLE
PS> $NC=Show-ColorPicker -SortMode BrightnessDescending -CurrentColor ([System.Drawing.Color]::Yellow) -HT
This example displays a ColorPicker with colors sorted by Brightness and Yellow loaded as the current color.
#>
function Show-ColorPicker{
	param (
		[Parameter()][Alias('S')][ColorPickerSortMode]$SortMode = [ColorPickerSortMode]::BrightnessDescending,
		[Parameter()][Alias('C')][System.Drawing.Color]$CurrentColor = [System.Drawing.Color]::FromKnownColor([System.Drawing.KnownColor]::Control),
		[Parameter()][Alias('HT')][Switch]$HideTransparent,
		[Parameter()][Alias('RC')][Switch]$ReturnColorInfo)

	Add-Type -A System.Windows.Forms

	$Script:NCV = $null #Initialize Return Value
	$CSTMask = "Name: {0},`r`nHex Value: {1:X8}"
	$ImageRes = $Env:SystemRoot+'\System32\imageres.dll'

	#region Form Objects
	#Add objects for Form
	$ToolTipProvider = [Windows.Forms.ToolTip]::New()
	$FrmColorPicker = [Windows.Forms.Form]::New()
	$PnlColorGrid = [Windows.Forms.Panel]::New()
	$GrpSort = [Windows.Forms.GroupBox]::New()
	$LbxSort = [Windows.Forms.ListBox]::New()
	$ColorButtons = @(for ($C = 1; $C -le $ArlColors.Count+1; $C++) {[Windows.Forms.Button]::New()})
	$Labels = @(for ($C = 1; $C -le 4; $C++) {[Windows.Forms.Label]::New()})
	$Buttons = @(for ($C = 1; $C -le 2; $C++) {[Windows.Forms.Button]::New()})
	#endregion

	#region Initialize ToolTipProvider
	$ToolTipProvider.AutoPopDelay = 5000
	$ToolTipProvider.InitialDelay = 100
	$ToolTipProvider.ReshowDelay = 100
	#endregion

	#region Custom Event Code Blocks
	<#
	define a scriptblock to display the tooltip
	add a _MouseHover event to display the corresponding tool tip
	e.g. $txtPath.add_MouseHover($ShowHelp)
	#>
	$ShowColorToolTip={
	#display popup help
	#each value is the name of a control on the panel.
	$ToolTipProvider.SetToolTip($This,$This.BackColor.Name)
	}
	#Selection Button Click Events
	$BtnOk_Click = {
		$Script:NCV = Set-ReturnValue -ExitButton OK -NewColor $Labels[[FormLabel]::NewColorPatch].BackColor
		$FrmColorPicker.Close()
	}
	$BtnCancel_Click = {
		$Script:NCV = Set-ReturnValue -ExitButton Cancel
		$FrmColorPicker.Close()
	}
	#ColorGrid Button Click
	$ColorButton_Click = {
		$Labels[[FormLabel]::NewColorPatch].BackColor = $This.Backcolor
		$Labels[[FormLabel]::NewColorPatch].Text = $CSTMask -f $This.Backcolor.Name,$This.Backcolor.ToArgb()
		Set-TextColor -Control $This -LabelIndex ([FormLabel]::NewColorPatch)
	}
	#SortIndex Changed
	$SortIndexChanged = {
		Sort-Colors -SortMode $LbxSort.SelectedIndex
		Set-GridColors
	}
	#endregion
	
	#region ColorPicker Form
	$FrmColorPicker.Name = 'FrmColorPicker'
	$FrmColorPicker.ClientSize = [Drawing.Size]::New(365, 355)
	$FrmColorPicker.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
	$FrmColorPicker.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$FrmColorPicker.Text = 'Blueflame Dynamics - PSColorPicker' 
	$FrmColorPicker.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($ImageRes, 186, 32)
	$FrmColorPicker.MaximizeBox = `
	$FrmColorPicker.MinimizeBox = $False
	$Form_BringToTop = {
		$This.TopMost = $True
		$This.BringToFront()
		$This.TopMost = $False
	}
	$FrmColorPicker.Add_Shown($Form_BringToTop)
	#endregion

	#region ColorPanel
	$PnlColorGrid.Name = 'PnlColorGrid'
	$PnlColorGrid.Size = [Drawing.Size]::New(143, 313)
	$PnlColorGrid.Location = [Drawing.Point]::New(13, 13)
	$PnlColorGrid.Anchor = Get-Anchor -T -L -B -R
	$PnlColorGrid.BackColor = [System.Drawing.SystemColors]::Control
	$PnlColorGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$PnlColorGrid.Cursor = Get-Cursor -Mode Default
	$PnlColorGrid.Visible = `
	$PnlColorGrid.Enabled = $True
	$PnlColorGrid.TabStop = $False
	$FrmColorPicker.Controls.Add($PnlColorGrid)
	#endregion

	#region Panel Buttons
	$C = 1 #Column
	$R = 1 #Row
	$S = 1 #Sequence
	$X = 3 #Horizontal Coordinate
	$Y = 3 #Vertical Coordinate
	For($Idx = 0; $Idx -lt $ColorButtons.Count; $Idx++){
		$ColorButtons[$Idx].Name = "BtnColorGrid{0:000}" -f $S
		$ColorButtons[$Idx].Parent = $PnlColorGrid
		$ColorButtons[$Idx].Size = [Drawing.Size]::New(15, 15)
		$ColorButtons[$Idx].Location = [Drawing.Point]::New($X, $Y)
		$ColorButtons[$Idx].FlatStyle = [Windows.Forms.FlatStyle]::Popup
		$ColorButtons[$Idx].Visible = `
		$ColorButtons[$Idx].Enabled = `
		$ColorButtons[$Idx].TabStop = $True
		$ColorButtons[$Idx].TabIndex = $Idx
		#Anchor Tranparent on Button 1 of the grid
		If($Idx -eq 0)
			{$ColorButtons[$Idx].Backcolor = [System.Drawing.Color]::Transparent}
		Else
			{$ColorButtons[$Idx].Backcolor = [System.Drawing.Color]::FromName($ArlColors[$Idx-1].Name)}
		$ColorButtons[$Idx].Add_MouseHover($ShowColorToolTip)
		$ColorButtons[$Idx].Add_Click($ColorButton_Click)
		#Layout buttons in an 8 button wide grid
		If($S % 8 -eq 0){
			$R++
			$C = 1
			$X = 3
			$Y += 17}
		Else{
			$C++
			$X += 17
		}
		$S++
	}
	$ColorButtons[0].Enabled = $ColorButtons[0].Visible = !$HideTransparent.IsPresent
	#endregion

	#region Labels
	$LblText = @('Current Color','Undefined','New Color','Undefined')
	$LblPos  = @(0,0,0,0)
	$LblSize = @(0,0,0,0)
	$LblPos[[FormLabel]::OldColorLabel] = [Drawing.Point]::New(165,13)
	$LblPos[[FormLabel]::OldColorPatch] = [Drawing.Point]::New(165,38)
	$LblPos[[FormLabel]::NewColorLabel] = [Drawing.Point]::New(165,90)
	$LblPos[[FormLabel]::NewColorPatch] = [Drawing.Point]::New(165,115)
	$LblSize[[FormLabel]::OldColorLabel] = [Drawing.Size]::New(100,23)
	$LblSize[[FormLabel]::OldColorPatch] = [Drawing.Size]::New(170,50)
	$LblSize[[FormLabel]::NewColorLabel] = [Drawing.Size]::New(100,23)
	$LblSize[[FormLabel]::NewColorPatch] = [Drawing.Size]::New(170,50)
	For($Idx = 0; $Idx -lt $Labels.Count; $Idx++){
		$Labels[$Idx].Name = "Label{0}" -f $Idx+1
		$Labels[$Idx].Parent = $FrmColorPicker
		$Labels[$Idx].Visible = `
		$Labels[$Idx].Enabled = $True
		$Labels[$Idx].Location = $LblPos[$Idx]
		$Labels[$Idx].TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$Labels[$Idx].BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
		$Labels[$Idx].Text = $LblText[$Idx]
		$Labels[$Idx].Size = $LblSize[$Idx]
	}
	$Labels[[FormLabel]::OldColorPatch].Backcolor = $CurrentColor
	$Labels[[FormLabel]::OldColorPatch].Text = $CSTMask -f $CurrentColor.Name,$CurrentColor.ToArgb()
	Set-TextColor -Control $Labels[[FormLabel]::OldColorPatch] -LabelIndex ([FormLabel]::OldColorPatch)
	#endregion

	#region GroupBox
	$GrpSort.Name = 'GrpSortBy'
	$GrpSort.Parent = $FrmColorPicker
	$GrpSort.Visible = `
	$GrpSort.Enabled = $True
	$GrpSort.Text = 'Sort By:'
	$GrpSort.Size = [Drawing.Size]::New(190, 142) 
	$GrpSort.Location = [Drawing.Point]::New(165,165)
	#endregion

	#region ListBox
	$LbxSort.Name = 'LbxSort'
	$LbxSort.Parent = $GrpSort
	$LbxSort.Visible = `
	$LbxSort.Enabled = `
	$LbxSort.Sorted = $True
	$LbxSort.Location = [Drawing.Point]::New(5,15)
	$LbxSort.Size = [Drawing.Size]::New(180, 121)

	$Modes = @()
	#Split Enum Names at Uppercase letters
	[ColorPickerSortMode].GetEnumNames()|ForEach-Object{
		$Parts = [Regex]::Matches($_, '[A-Z][a-z]+')
		Switch($Parts.Count){
			2 {$Mode = '{0} - {1}' -f $Parts[0], $Parts[1]}
			3 {$Mode = '{0} {1} - {2}' -f $Parts[0], $Parts[1], $Parts[2]}
			4 {$Mode = '{0} {1} {2} - {3}' -f $Parts[0], $Parts[1], $Parts[2], $Parts[3]}
			}
		If($Mode.StartsWith('Argb')){$Mode=$Mode.Replace('Argb','ARGB')}
		$Modes += $Mode
		}
	$LbxSort.SelectionMode = [System.Windows.Forms.SelectionMode]::One
	$LbxSort.Items.AddRange($Modes)
	$LbxSort.SelectedIndex = $SortMode
	$LbxSort.Add_SelectedIndexChanged($SortIndexChanged)
	#Perform Initial Sort, If needed.
	If($SortMode -ne [ColorPickerSortMode]::NameAscending)
		{Invoke-Command -ScriptBlock $SortIndexChanged} 
	#endregion

	#region Selection Buttons
	$BtnText = @('Ok','Cancel')
	$BtnPos = @(0,0)
	$BtnPos[0] = [Drawing.Point]::New(205,315)
	$BtnPos[1] = [Drawing.Point]::New(280,315)
	For($Idx = 0; $Idx -lt $Buttons.Count; $Idx++){
		$Buttons[$Idx].Name = "Btn{0}" -f $BtnText[$Idx]
		$Buttons[$Idx].Parent = $FrmColorPicker
		$Buttons[$Idx].Visible = `
		$Buttons[$Idx].Enabled = $True
		$Buttons[$Idx].Text = $BtnText[$Idx]
		$Buttons[$Idx].Location = $BtnPos[$Idx]
		$Buttons[$Idx].Size = [Drawing.Size]::New(75,30)
	}
	$Buttons[0].Add_Click($BtnOK_Click)
	$Buttons[1].Add_Click($BtnCancel_Click)
	#endregion

	[Void]$FrmColorPicker.ShowDialog()
	$Script:NCV
}

function Set-ReturnValue{
	param (
		[Parameter(Mandatory)][Alias('EB')][System.Windows.Forms.DialogResult]$ExitButton,
		[Parameter()][Alias('NC')][System.Drawing.Color]$NewColor = $Null)
	if($ReturnColorInfo.IsPresent -and $Null -ne $NewColor){
		[PSCustomObject][Ordered] @{ExitButton = $ExitButton;Color = $ArlColors.FindColorInfo($NewColor)}
	}Else{
		[PSCustomObject][Ordered] @{ExitButton = $ExitButton;Color = $NewColor}
	}
}

function Sort-Colors{
	param ([Parameter(Mandatory)][Alias('S')][ColorPickerSortMode]$SortMode)
	$SM = [ColorPickerSortMode] #Shortcut
	$Sort = {
	param(
		[Parameter(Mandatory)][Alias('P')][String]$Property,
		[Parameter()][Alias('D')][Switch]$Descending)
	#Sort Minor Key: Name Ascending, Requested Major Key. Preserves Input ArrayList.
	$RV = $ArlColors|Sort-Object -Property Name|Sort-Object -Property $Property -Descending:$Descending
	$ArlColors.Clear()
	$ArlColors.AddRange($RV)
	}
	Switch($SortMode){
	$SM::ArgbHexValueAscending  {& $Sort -Property ARGB}
	$SM::ArgbHexValueDescending {& $Sort -Property ARGB -Descending}
	$SM::BlueChannelAscending   {& $Sort -Property B}
	$SM::BlueChannelDescending  {& $Sort -Property B -Descending}
	$SM::BrightnessAscending	{& $Sort -Property Brightness}
	$SM::BrightnessDescending   {& $Sort -Property Brightness -Descending}
	$SM::GreenChannelAscending  {& $Sort -Property G}
	$SM::GreenChannelDescending {& $Sort -Property G -Descending}
	$SM::HueAscending			{& $Sort -Property Hue}
	$SM::HueDescending			{& $Sort -Property Hue -Descending}
	$SM::NameAscending			{& $Sort -Property Name}
	$SM::NameDescending			{& $Sort -Property Name -Descending}
	$SM::RedChannelAscending	{& $Sort -Property R}
	$SM::RedChannelDescending   {& $Sort -Property R -Descending}
	$SM::SaturationAscending	{& $Sort -Property Saturation}
	$SM::SaturationDescending   {& $Sort -Property Saturation -Descending}
	}
}

function Set-GridColors{
	for($Idx = 0; $Idx -lt $ColorButtons.Count; $Idx++){
		#Anchor Tranparent on Button 1 of the grid
		if($Idx -eq 0)
			{$ColorButtons[$Idx].Backcolor = [System.Drawing.Color]::Transparent}
		else
			{$ColorButtons[$Idx].Backcolor = [System.Drawing.Color]::FromName($Script:ArlColors[$Idx-1].Name)}
	}
}

function Set-TextColor{
	param(
		[Parameter(Mandatory)][Alias('C')][Windows.Forms.Control]$Control,
		[Parameter(Mandatory)][Alias('I')][Int]$LabelIndex)
	
	if($Control.Backcolor.GetBrightness() -le 0.49) 
		{$Labels[$LabelIndex].ForeColor = [System.Drawing.Color]::White}
	else
		{$Labels[$LabelIndex].ForeColor = [System.Drawing.Color]::Black}
	#For Blue\Black Color Blind
	If($Control.Backcolor.Name -eq "Blue"){$Labels[$LabelIndex].ForeColor = [System.Drawing.Color]::White}
}