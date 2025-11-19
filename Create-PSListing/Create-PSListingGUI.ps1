<#
.NOTES
	File Name:	Create-PSListingGUI.ps1
	Version:	1.1 - 11/18/2025
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	03/14/2024
	History:
		V1.1 - 11/18/2025 - Added Sort Options
		V1.0 - 03/14/2024 - Original Release

.SYNOPSIS
	This script provides a GUI front-end for my Create-PSListing.ps1 V2.0+ script
	to allow it to be called in a prompted manner.  

.DESCRIPTION
	Calls the Create-PSListing V2.0+ script in a prompted manner or displays diagnostics.
	Create-PSListing.ps1 V2.0 added direct support for a Powershell Token report similar
	to a cross-reference in other programming languages.  Most of the dialog settings may be 
	saved on exit to a Create-PSListingGUI.json when the associated checkbox has been checked
	and the dialog is exited by use of the Ok button, unless the Diagnostic checkbox is checked.
	The input file, Save-On Exit, Diagnostic, & Token Report Only values are NOT saved.
	Due to what appears to be a scope issue, when run in Powershell Core 6+ an 
	Add-Type -AssemblyName System.Windows.Forms must be performed first.

-----------------------------[Special Note]-------------------------
You may register this in your VSCode_profile using:
Register-EditorCommand -Name 'BlueflameDynamics.Build-PSListing' `
-DisplayName 'Create PS Listing' `
-ScriptBlock {.\Create-PSListingGUI.ps1} -SuppressOutput
--------------------------------------------------------------------

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------	

.PARAMETER Path
	Fully Qualified Input Script Name.
.PARAMETER StartPosition
	Dialog StartPosition of CenterParent, CenterScreen (Default), CenterTop (Center screen, Top of ISE panel),
	or Point. When Point is selected two dynamic parameters X & Y become available. These
	parameters are the X,Y coordinates for the onscreen dialog.	
.EXAMPLE
	PS> .\Create-PSListingGUI.ps1 
	Runs this script displaying the GUI interface for Create-PSListing.ps1 V2.0+
.EXAMPLE
	PS> .\Create-PSListingGUI.ps1 -StartPosition Point -X 250 -Y 97
	Runs this script displaying the GUI Dialog at screen point: 250,97
.EXAMPLE
	PS> '.\AES-Email.ps1'|.\Create-PSListingGUI.ps1 -StartPosition Point -X 250 -Y 97
	Runs this script displaying the GUI Dialog at screen point: 250,97 with the named input file.
#>
[CmdletBinding()]
param(
	[Parameter(
		ValueFromPipeline,
		HelpMessage = 'The path of the input file')][Alias('Path')][String]$FullName,
	[Parameter(HelpMessage = 'Dialog Window StartPosition')][Alias('WS')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('CenterParent','CenterScreen','CenterTop','Point')]
		[String]$StartPosition = 'CenterScreen')
DynamicParam{
# Set up the Run-Time Parameter Dictionary
$RuntimeParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::New()

if($StartPosition -eq 'Point'){
	#region X
	$Key = 'X'
	$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::New()
	$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::New()
	$ParameterAttribute.Mandatory = $True
	$ParameterAttribute.Position = 2
	$ParameterAttribute.HelpMessage = 'X-Coordinate'
	$AttributeCollection.Add($ParameterAttribute)
	$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::New($Key,[Int32],$AttributeCollection)
	$RuntimeParameterDictionary.Add($Key, $RuntimeParameter)
	#endregion
	#region Y
	$Key = 'Y'
	$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::New()
	$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::New()
	$ParameterAttribute.Mandatory = $True
	$ParameterAttribute.Position = 3
	$ParameterAttribute.HelpMessage = 'Y-Coordinate'
	$AttributeCollection.Add($ParameterAttribute)
	$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::New($Key,[Int32],$AttributeCollection)
	$RuntimeParameterDictionary.Add($Key, $RuntimeParameter)
	#endregion
	}
# When done building dynamic parameters, return
return $RuntimeParameterDictionary
}
Begin{
#region Add Types
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Path .\BlueflameDynamics.IconTools.dll
#endregion
#region Import Modules
Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\OSInfoLib.ps1 -Force
Import-Module -Name .\WinFormsLibrary.ps1 -Force
#endregion
#region Enums
Enum StartPosition{
	CenterParent
	CenterScreen
	CenterTop
	Point}
Enum FormButtonType{
	Ok
	Cancel
	Host}
Enum OptionRadioButtonType{
	Host
	Grid}
Enum TokenReportOutput{
	File
	GridView
	Host
	Html
	Raw}
Enum HostID{
	psISEHost
	VSCodeHost}
#endregion
#region Supported Host Editors
$HostEditors = @{
	'Windows PowerShell ISE Host' = [HostID]::psISEHost.ToString()
	'Visual Studio Code Host' = [HostID]::VSCodeHost.ToString()}
#endregion
#region Script Parameter Variables
$P = ($MyInvocation.MyCommand).Parameters
$V = $P['StartPosition'].Attributes.ValidValues
$My = [PSCustomObject][Ordered]@{
	PSTypeName = 'My.ScriptParameters'
	Title = 'Create-PSListing GUI'
	Name = ($MyInvocation.MyCommand).Name
	SettingsFile = $MyInvocation.InvocationName.Replace('.ps1','.json')
	Params = $P
	HostId = $HostEditors[$Host.Name]
	Icon = [PSCustomObject][Ordered]@{
		File = $Env:SystemRoot+'\System32\ImageRes.dll'
		Index = If((Get-OSVersion).IsWin11){312}Else{311}}
	Location = [PSCustomObject][Ordered]@{
		Values = $V
		Index = [Array]::IndexOf($V,$StartPosition)
		Point = $Null}
}
If($My.Location.Index -eq [StartPosition]::Point){
	$My.Location.Point = [Drawing.Point]::New(
		$RuntimeParameterDictionary.X.Value,
		$RuntimeParameterDictionary.Y.Value)
}
Remove-Variable -Name P,V
#endregion
#region Utility Functions
Function Export-Settings{
    param(
        [Parameter(Mandatory)][PSTypeName('My.DialogResult')]$DR,
        [Parameter(Mandatory)][string]$File)
    if ($DR.Diagnostics) { return }
    #Remove Properties Excluded from Export
    foreach ($prop in 'SaveOnExit','Diagnostics','SelectedFile'){
        if ($DR.PSObject.Properties.Match($prop)) {$DR.PSObject.Properties.Remove($prop)}
    }

    if ($DR.PSObject.Properties.Match('TokenReport')) {
        $tokenReport = $DR.TokenReport
        if ($tokenReport -and $tokenReport.PSObject.Properties.Match('TokenReportOnly')) {
            $tokenReport.PSObject.Properties.Remove('TokenReportOnly')
        }
    }

    #Set Object TypeName
    $newType = 'My.Settings'
    if ($DR.PSTypeNames.Count -gt 0) {
        $DR.PSTypeNames[0] = $newType
    } else {
        $DR.PSTypeNames.Insert(0,$newType)
    }

    Add-Member -InputObject $DR -MemberType NoteProperty -Name TypeName -Value $newType -Force

    #Export to JSON
    try {
        $DR | ConvertTo-Json -Depth $DR.JsonDepth | Set-Content -Path $File -ErrorAction Stop
    } catch {
        Write-Error "Failed to export settings: $_"
    }
}
Function Import-Settings{
	param([Parameter(Mandatory)][String]$File)
	If(!(Test-Exists -Mode File -Location $File)){Return $Null}
	$RV = Get-Content -Path $File|ConvertFrom-Json
	$RV.PSTypeNames.Insert(0,$RV.TypeName)
	$RV
}
#endregion
#region Custom Classes
Class Settings{
	$ChkBxText = @('Include Tokens','Open In Web Browser','Output To Temp','Save on Exit')
	$LabelText = @('HTML Font Size','Tab Expansion Width','Listing Line # Width')
	$IV	= @(1,2,1)
	$Min = @(10,2,2)
	$Max = @(24,8,8)
	$Value = @(16,4,4)
	$GroupBox = (New-ObjectArray -TypeName Windows.Forms.GroupBox -Count 1)
	$Labels = (New-ObjectArray -TypeName Windows.Forms.Label -Count $This.LabelText.Count)
	$NumUpDowns = (New-ObjectArray -TypeName Windows.Forms.NumericUpDown -Count $This.LabelText.Count)
	$CheckBoxes = (New-ObjectArray -TypeName Windows.Forms.Checkbox -Count $This.ChkBxText.Count)

	Settings(){
		$This.GroupBox.Name = 'GrpSettings'
		$This.GroupBox.Text = 'Settings'
		$This.GroupBox.Size = [Drawing.Size]::New(203,190)
		$This.GroupBox.Location = [Drawing.Point]::New(10,3)
		$This.GroupBox.Controls.AddRange($This.Labels)
		$This.GroupBox.Controls.AddRange($This.NumUpDowns)
		$This.GroupBox.Controls.AddRange($This.CheckBoxes)
		#region Labels
		$C = 0; $Y = 23
		ForEach($Lbl in $This.Labels){
			$Lbl.Size = [Drawing.Size]::New(115,15)
			$Lbl.Location = [Drawing.Point]::New(10,$Y)
			$Lbl.Name = 'Lbl'+$This.LabelText[$C].Replace(' ','')
			$Lbl.TabStop = $False
			$Lbl.Text = $This.LabelText[$C]
			$Y+=20
			$C++
		}
		#endregion
		#region NumUpDowns
		$C = 0; $Y = 19; $T = 3
		ForEach($Nud in $This.NumUpDowns){
			$Nud.Size = [Drawing.Size]::New(49,23)
			$Nud.Location = [Drawing.Point]::New(140,$Y)
			$Nud.Name = 'Nud'+$This.LabelText[$C].Replace(' ','')
			$Nud.TabStop = $True
			$Nud.TabIndex = $T
			$Nud.Increment = $This.IV[$C]
			$Nud.Maximum = $This.Max[$C]
			$Nud.Minimum = $This.Min[$C]
			$Nud.Value = $This.Value[$C]
			$Y+=19
			$C++
			$T++
		}
		#endregion
		#region Checkboxes
		$C = 0; $Y = 87; #$T = 6
		ForEach($CB in $This.CheckBoxes){
			$CB.Name = 'Chk'+$This.ChkBxText[$C].Replace(' ','')
			$CB.Text = $This.ChkBxText[$C]
			$CB.TabStop = `
			$CB.AutoSize = $True
			$CB.TabIndex = $T
			$CB.Location = [Drawing.Point]::New(14,$Y)
			$CB.Size = [Drawing.Size]::New(140,19)
			$CB.UseVisualStyleBackColor = $True
			$Y+=20
			$C++
			$T++
		}
		#endregion
	}
}
Class Options{
	$DiagOuts = @('Host','Grid')
	$GroupBox = [Windows.Forms.GroupBox]::New()
	$ChkDiag = [Windows.Forms.Checkbox]::New()
	$TblLayoutPanel = [Windows.Forms.TableLayoutPanel]::New() 
	$RadioButtons = (New-ObjectArray -TypeName Windows.Forms.RadioButton -Count $This.DiagOuts.Count)

	Options(){
		#region GroupBox
		$This.GroupBox.Name = 'GrpOptions'
		$This.GroupBox.Text = 'Options'
		$This.GroupBox.Size = [Drawing.Size]::New(314,45)
		$This.GroupBox.Location = [Drawing.Point]::New(219,3)
		$This.GroupBox.TabStop = $False
		$This.GroupBox.Controls.Add($This.ChkDiag)
		$This.GroupBox.Controls.Add($This.TblLayoutPanel)
		$This.TblLayoutPanel.Controls.AddRange($This.RadioButtons)
		#endregion
		#region ChkDiag
		$This.ChkDiag.Location = [Drawing.Point]::New(14,18)
		$This.ChkDiag.Size = [Drawing.Size]::New(88,19)
		$This.ChkDiag.Name = 'ChkDiag'
		$This.ChkDiag.Text = 'Diagnostics'
		$This.ChkDiag.UseVisualStyleBackColor = `
		$This.ChkDiag.TabStop = $True
		$This.ChkDiag.TabIndex = 9

		#endregion
		#region TableLayoutPanel
		$This.TblLayoutPanel.Name = 'TlpButtons'
		$This.TblLayoutPanel.Size = [Drawing.Size]::New(114,26)
		$This.TblLayoutPanel.Location = [Drawing.Point]::New(110,14)
		$This.TblLayoutPanel.Anchor = Get-Anchor -T -L
		$This.TblLayoutPanel.ColumnCount = 2
		$This.TblLayoutPanel.RowCount = 1
		For($C=0;$C -lt $This.TblLayoutPanel.ColumnCount;$C++){
			[Void]$This.TblLayoutPanel.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::New(
				[System.Windows.Forms.SizeType]::Percent,50))}
		[Void]$This.TblLayoutPanel.RowStyles.Add([System.Windows.Forms.RowStyle]::New(
			[System.Windows.Forms.SizeType]::Percent,50))
		#endregion
		#region DiagRadioButtons
		$C = 0; $X = 105; $T = 10
		ForEach($RB in $This.RadioButtons){
			$RB.Location = [Drawing.Point]::New($X,16)
			$RB.Size = [Drawing.Size]::New(50,19)
			$RB.Text = $This.DiagOuts[$C]
			$RB.Name = 'Rdo'+$This.DiagOuts[$C]
			$RB.TabStop = $True
			$RB.TabIndex = $T
			$X+=50
			$C++
			$T++
		}
		$This.RadioButtons[[OptionRadioButtonType]::Host].Checked = $True
		#endregion
	}
}
Class TokenRpt{
	$ChkBxText = @('Token Report Only','Include Comment Type','Include Start Property','Sort By Type','Sort By Content')
	$RBText = @('File','GridView','Host','Html','Raw')
	$GroupBox = [Windows.Forms.GroupBox]::New()
	$Label = [Windows.Forms.Label]::New()
	$MaxCommentLength = [Windows.Forms.NumericUpDown]::New()
	$RadioButtons = (New-ObjectArray -TypeName Windows.Forms.RadioButton -Count $This.RBText.Count)
	$CheckBoxes = (New-ObjectArray -TypeName Windows.Forms.Checkbox -Count $This.ChkBxText.Count)

	TokenRpt(){
		#region GroupBox
		$This.GroupBox.Name = 'GrpTokenRpt'
		$This.GroupBox.Text = 'Token Report Options'
		$This.GroupBox.Size = [Drawing.Size]::New(314,145)
		$This.GroupBox.Location = [Drawing.Point]::New(219,48)
		$This.GroupBox.Enabled = `
		$This.GroupBox.TabStop = $False
		$This.GroupBox.Controls.Add($This.Label)
		$This.GroupBox.Controls.Add($This.MaxCommentLength)
		$This.GroupBox.Controls.AddRange($This.RadioButtons)
		$This.GroupBox.Controls.AddRange($This.CheckBoxes)
		#endregion
		#region Labels
		$This.Label.Name = 'LblMaxCommentLength'
		$This.Label.Text = 'Max Comment Length:'
		$This.Label.Size = [Drawing.Size]::New(137,20)
		$This.Label.Location = [Drawing.Point]::New(101,($This.ChkBxText.Count * 23) + 8)
		#endregion  
 		#region NumUpDowns
 		$This.MaxCommentLength.Name = 'MaxCommentLength'
  		$This.MaxCommentLength.Size  = [Drawing.Size]::New(50,23)
		$This.MaxCommentLength.Location = [Drawing.Point]::New(244,($This.ChkBxText.Count * 23) + 5)
		$This.MaxCommentLength.Maximum = 100
		$This.MaxCommentLength.Minimum = 10
		$This.MaxCommentLength.Increment = 1
		$This.MaxCommentLength.Value = 50
		$This.MaxCommentLength.TabStop = $True
		$This.MaxCommentLength.TabIndex = 19
		#endregion 
		#region RadioButtons
		$C = 0; $Y = 20; $T = 12
		ForEach($RB in $This.RadioButtons){
			$RB.Name = 'Rdo'+$This.RBText[$C].Replace(' ','')
			$RB.Text = $This.RBText[$C]
			$RB.Location = [Drawing.Point]::New(14,$Y)
			$RB.Size = [Drawing.Size]::New(140,19)
			$RB.TabIndex = $T
			$RB.TabStop = `
			$RB.UseVisualStyleBackColor = `
			$RB.AutoSize = $True
			$Y+=20
			$C++
			$T++
		}
		$This.RadioButtons[[TokenReportOutput]::File].Checked = $True
		#endregion
		#region Checkboxes
		$C = 0; $Y = 20; #$T = 17
		ForEach($CB in $This.CheckBoxes){
			$CB.Name = 'Chk'+$This.ChkBxText[$C].Replace(' ','')
			$CB.Text = $This.ChkBxText[$C]
			$CB.Location = [Drawing.Point]::New(100,$Y)
			$CB.Size = [Drawing.Size]::New(140,19)
			$CB.TabIndex = $T
			$CB.TabStop = `
			$CB.UseVisualStyleBackColor = `
			$CB.AutoSize = $True
			$Y+=20
			$C++
			$T++
		}
		#endregion
	}
}
Class SelectedFile{
	$GroupBox = [Windows.Forms.GroupBox]::New()
	$Label = [Windows.Forms.Label]::New()
	$TextBox = [Windows.Forms.TextBox]::New()
	$Button = [Windows.Forms.Button]::New()

	SelectedFile(){
		#region GroupBox
		$This.GroupBox.Name = 'GrpSelectedFile'
		$This.GroupBox.Text = 'Selected File'
		$This.GroupBox.Size = [Drawing.Size]::New(524,66)
		$This.GroupBox.Location = [Drawing.Point]::New(10,192)
		$This.GroupBox.Controls.Add($This.Label)
		$This.GroupBox.Controls.Add($This.TextBox)
		$This.GroupBox.Controls.Add($This.Button)
		#endregion
		#region Label
		$This.Label.Name = 'LblFileIn'
		$This.Label.Text = 'File:'
		$This.Label.Size = [Drawing.Size]::New(28,15)
		$This.Label.Location = [Drawing.Point]::New(10,32)
		#endregion
		#region TextBox
		$This.TextBox.Name = 'TxtSelectedFile'
		$This.TextBox.Size = [Drawing.Size]::New(440,23)
		$This.TextBox.Location = [Drawing.Point]::New(40,28)
		$This.TextBox.TabStop = $True
		$This.TextBox.TabIndex = 21
		#endregion
		#region Button
		$This.Button.Name = 'BtnOpenFile'
		$This.Button.Text = '>>'
		$This.Button.Size = [Drawing.Size]::New(32,23)
		$This.Button.Location = [Drawing.Point]::New(486,27)
		$This.Button.TabStop = $True
		$This.Button.TabIndex = 22
		#endregion
	}
}
Class MainForm{ 
	$BtnText = @('Ok','Cancel','HostInfo')
	$Form = [Windows.Forms.Form]::New()
	$Panel = [Windows.Forms.Panel]::New()
	$TblLayoutPanel = [Windows.Forms.TableLayoutPanel]::New()
	$Buttons = (New-ObjectArray -TypeName Windows.Forms.Button -Count $This.BtnText.Count)

	MainForm(){
		#region Form
		$This.Form.Name = 'FrmMain'
		$This.Form.Text = $Script:My.Title + ' Dialog'
		$This.Form.Size = [Drawing.Size]::New(580,355)
		$This.Form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
		$This.Form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
		$This.Form.MaximizeBox = `
		$This.Form.MinimizeBox = $False
		#endregion
		#region Panel 
		$This.Panel.Name = 'Panel'
		$This.Panel.Location = [Drawing.Point]::New(8,8)
		$This.Panel.Size = [Drawing.Size]::New(546,265)
		$This.Panel.BackColor = [Drawing.SystemColors]::ControlDark
		$This.Form.Controls.Add($This.Panel)
		#endregion 
		#region TableLayoutPanel
		$This.TblLayoutPanel.Name = 'TlpButtons'
		$This.TblLayoutPanel.Size = [Drawing.Size]::New(209, 31)
		$This.TblLayoutPanel.Location = [Drawing.Point]::New(334,273)
		$This.TblLayoutPanel.Anchor = Get-Anchor -B -R
		$This.TblLayoutPanel.ColumnCount = 3
		$This.TblLayoutPanel.RowCount = 1
		For($C=0;$C -lt $This.TblLayoutPanel.ColumnCount;$C++){
			[Void]$This.TblLayoutPanel.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::New(
			[System.Windows.Forms.SizeType]::Percent,33))}
		[Void]$This.TblLayoutPanel.RowStyles.Add([System.Windows.Forms.RowStyle]::New(
			[System.Windows.Forms.SizeType]::Percent,50))
		#endregion
		#region Buttons
		$C = 0
		ForEach($Btn in $This.Buttons){
			$Btn.Text = '&'+$This.BtnText[$C]
			$Btn.Name = 'Btn'+$This.BtnText[$C]
			$Btn.TabStop = $True
			$Btn.TabIndex = $C
			$C++
		}

		$This.Buttons[[FormButtonType]::Ok].DialogResult = [Windows.Forms.DialogResult]::OK
		$This.Buttons[[FormButtonType]::Cancel].DialogResult = [Windows.Forms.DialogResult]::Cancel

		$This.Form.AcceptButton = $This.Buttons[[FormButtonType]::Ok]
		$This.Form.CancelButton = $This.Buttons[[FormButtonType]::Cancel]
		$This.TblLayoutPanel.Controls.AddRange($This.Buttons)
		$This.Buttons[[FormButtonType]::Host].Add_Click({Show-HostInfo -Owner $FrmMain.Form -App $FrmMain.Form.Text})
		#endregion
		#region Add Some Controls
		$This.Panel.Controls.Add(([Settings]::New()).GroupBox)
		$This.Panel.Controls.Add(([Options]::New()).GroupBox)
		$This.Panel.Controls.Add(([TokenRpt]::New()).GroupBox)
		$This.Panel.Controls.Add(([SelectedFile]::New()).GroupBox)
		$This.Form.Controls.Add($This.TblLayoutPanel)
		#endregion
		#region Textbox Validation event
		$ValidateText = {
			If($This.Text.Length -ne 0 -and !([IO.FileInfo]::New($This.Text)).Exists){
				Show-MessageBoxEx `
					-Messsage 'Valid input ps1\psm1 file required!'`
					-Title 'Invalid Input File'`
					-I ([System.Windows.Forms.MessageBoxIcon]::Warning)
			}
		}
		$This.Panel.Controls.
			Item('GrpSelectedFile').Controls.
			Item('TxtSelectedFile').Add_Validating($ValidateText)
		#endregion
		#region Open button click event
		$OpenFileClick = {
			$RV = Invoke-OpenFileDialog `
				-ReturnMode Filename `
				-FileFilter ('{0}Scripts|*.ps1|{0}Modules|*.psm1' -f 'Powershell ')`
				-Title $Script:My.Title `
				-File $This.Parent.Controls.Item('TxtSelectedFile').Text
			If($RV -ne ''){$This.Parent.Controls.Item('TxtSelectedFile').Text = $RV}
		}
		$This.Panel.Controls.
			Item('GrpSelectedFile').Controls.
			Item('BtnOpenFile').Add_Click($OpenFileClick)
		#endregion
		#region ChkIncludeTokens CheckStateChanged event
		$This.Panel.Controls.
			Item('GrpSettings').Controls.
			Item('ChkIncludeTokens').Add_CheckStateChanged({
				$GTR = $This.Parent.Parent.Controls.Item('GrpTokenRpt')
				$GTR.Enabled = $This.Checked
				if(!$GTR.Enabled){$GTR.Controls.Item('ChkTokenReportOnly').Checked = $False}})
		#endregion
		#region Form Shown event
		$This.Form.Add_Shown({
			$This.TopMost = $True
			$This.BringToFront()
			$This.TopMost = $False
		})
		#endregion
	}
}
#endregion
}
Process{
#region Main Process
$FrmMain = [MainForm]::New()
$FrmMain.Form.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($My.Icon.File,$My.Icon.Index,16)
#region Groupbox Control Splats
$FPC = $FrmMain.Panel.Controls
$GrpSettings = $FPC.Item('GrpSettings').Controls
$GrpOptions = $FPC.Item('GrpOptions').Controls
$GrpTokenRpt = $FPC.Item('GrpTokenRpt').Controls
$GrpSelectedFile = $FPC.Item('GrpSelectedFile').Controls
#endregion
#region Local Functions
Function Mount-Settings{
	#Set Input File to current ISE\VSCode file, may be overridden
	$I = $GrpSelectedFile.Item('TxtSelectedFile')
	$I.Text = Switch([Int][HostID]::($My.HostID)){
		0	{$psISE.CurrentFile.FullPath;Break}
		1	{($PSEditor.GetEditorContext()).CurrentFile.Path;Break}
	}
	If($FullName.Length -gt 0){$I.Text = $FullName}

	Add-Member -InputObject $My -MemberType NoteProperty -Name Settings -Value (Import-Settings -File $My.SettingsFile)

	If($Null -ne $My.Settings){
		$GrpOptions.Item('TlpButtons').Controls.Item('Rdo'+$My.Settings.DiagOut).Checked = $True
		$GrpSettings.Item('NudHTMLFontSize').Value = $My.Settings.HTMLFontSize
		$GrpSettings.Item('NudTabExpansionWidth').Value = $My.Settings.TabExpansionWidth
		$GrpSettings.Item('NudListingLine#Width').Value = $My.Settings.ListingLineNoWidth
		$GrpSettings.Item('ChkIncludeTokens').Checked = $My.Settings.IncludeTokens
		$GrpSettings.Item('ChkOpenInWebBrowser').Checked = $My.Settings.OpenInWebBrowser
		$GrpSettings.Item('ChkOutputToTemp').Checked = $My.Settings.OutputToTemp
		$GrpTokenRpt.Item('ChkIncludeCommentType').Checked = $My.Settings.TokenReport.IncludeCommentType
		$GrpTokenRpt.Item('ChkIncludeStartProperty').Checked = $My.Settings.TokenReport.IncludeStartProperty
		$GrpTokenRpt.Item('ChkSortByType').Checked = $My.Settings.TokenReport.SortByType
		$GrpTokenRpt.Item('ChkSortByContent').Checked = $My.Settings.TokenReport.SortByContent
		$GrpTokenRpt.Item('MaxCommentLength').Value = $My.Settings.TokenReport.MaxCommentLength
		$GrpTokenRpt.Item('Rdo'+$My.Settings.TokenReport.OutputTo).Checked = $True
	}
}
Function Get-DialogResults{
	# Enforce OutputTo: HTML when detected Mode is Both
	if($GrpSettings.Item('ChkIncludeTokens').Checked -and
		!$GrpTokenRpt.Item('ChkTokenReportOnly').Checked){
			$GrpTokenRpt.Item('RdoHtml').Checked = $True}

	$ClickedRadioButton = $GrpTokenRpt|Where-Object -FilterScript {$_.GetType().Name -eq 'RadioButton' -and $_.Checked}
	$R = [PSCustomObject][Ordered]@{
		PSTypeName = 'My.DialogResult'
		JsonDepth = 2
		Mode = 'Listing'
		Diagnostics = $GrpOptions.Item('ChkDiag').Checked
		DiagOut = if($GrpOptions.Item('TlpButtons').Controls.Item('RdoGrid').Checked){'Grid'}Else{'Host'}
		HTMLFontSize = $GrpSettings.Item('NudHTMLFontSize').Value
		TabExpansionWidth = $GrpSettings.Item('NudTabExpansionWidth').Value
		ListingLineNoWidth = $GrpSettings.Item('NudListingLine#Width').Value
		IncludeTokens = $GrpSettings.Item('ChkIncludeTokens').Checked
		OpenInWebBrowser = $GrpSettings.Item('ChkOpenInWebBrowser').Checked
		OutputToTemp = $GrpSettings.Item('ChkOutputToTemp').Checked
		SaveOnExit = $GrpSettings.Item('ChkSaveonExit').Checked
		TokenReport = [PSCustomObject][Ordered]@{
			OutputTo = $ClickedRadioButton.Text
			TokenReportOnly = $GrpTokenRpt.Item('ChkTokenReportOnly').Checked
			IncludeCommentType = $GrpTokenRpt.Item('ChkIncludeCommentType').Checked
			IncludeStartProperty = $GrpTokenRpt.Item('ChkIncludeStartProperty').Checked
			SortByType = $GrpTokenRpt.Item('ChkSortByType').Checked
			SortByContent = $GrpTokenRpt.Item('ChkSortByContent').Checked
			MaxCommentLength = $GrpTokenRpt.Item('MaxCommentLength').Value
		}
		SelectedFile = $GrpSelectedFile.Item('TxtSelectedFile').Text
	}
	if($R.IncludeTokens){$R.Mode = 'Both'}
	if($R.TokenReport.TokenReportOnly){$R.Mode = 'Tokens'}
	Return $R
}
#endregion
Switch($My.Location.Index){
	([Int][StartPosition]::CenterParent)	{$FrmMain.Form.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent}
	([Int][StartPosition]::CenterTop)		{$FrmMain.Form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
		$FrmMain.Form.Location = [Drawing.Point]::New(([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width - $FrmMain.Form.Size.Width)/2,98)}
	([Int][StartPosition]::Point)			{$FrmMain.Form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
		$FrmMain.Form.Location = $My.Location.Point}
}
#region Load Settings, Show Dialog, & Get Results
Mount-Settings
if($FrmMain.Form.ShowDialog() -eq [Windows.Forms.DialogResult]::Cancel){Exit}
$DR = Get-DialogResults
#endregion
#region Diagnostic Output\Exit
if($DR.Diagnostics){
	if($Host.Name -ne 'Default Host'){Clear-Host} 
	If($GrpOptions.Item('TlpButtons').Controls.Item('RdoGrid').Checked){
		$DR|Out-GridView -Title 'Dialog Results' -Wait}
	Else{$DR|Out-Host}
	Exit
}
#endregion
#region Create-PSListing Call
.\Create-PSListing.ps1 `
	-Path $DR.SelectedFile`
	-Mode $DR.Mode`
	-FontSize $DR.HTMLFontSize`
	-TabWidth $DR.TabExpansionWidth`
	-LineNoWidth $DR.ListingLineNoWidth`
	-OutputTo $DR.TokenReport.OutputTo`
	-MaxContentLength $DR.TokenReport.MaxCommentLength`
	-ShowComments:$DR.TokenReport.IncludeCommentType`
	-ShowStart:$DR.TokenReport.IncludeStartProperty`
	-SortByType:$DR.TokenReport.SortByType`
	-SortByContent:$DR.TokenReport.SortByContent`
	-OutputToTemp:$DR.OutputToTemp`
	-OpenInWebBrowser:$DR.OpenInWebBrowser
#endregion
#endregion
}
End{If($DR.SaveOnExit){Export-Settings -DR $DR -File $My.SettingsFile}}