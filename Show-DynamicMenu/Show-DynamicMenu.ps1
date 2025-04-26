<#
.NOTES
Name:	 Show-DynamicMenu.ps1
Author:  Randy Turner
Version: 1.0
Date:	 11/16/2022

.SYNOPSIS
This script is designed to provide a WinForm Menu whose items are dynamically defined.
It may be called as a Cmdlet or Imported as a Module to allow direct access to it's functions.

.DESCRIPTION
This script is designed to provide a WinForm Menu whose items are dynamically defined.
It may be called as a Cmdlet or Imported as a Module to allow direct access to it's functions.
The return integer value is the index value of the selected item in the input array or
-1 for the Form Close Button(X), or -2 for the predefined menu Exit item\Alt-F4.

.PARAMETER MenuTitles
This is a String Array of the title text to appear in the menu.
This array is limited to 32767 items, surely more than needed.

.PARAMETER Title
This parameter is a String to appear in the form Title Bar.

.PARAMETER MenuTitle
This parameter is a String to appear in the menu dropdown Title.

.PARAMETER Icon
This parameter is an alternate Drawing.Icon to appear on the form.
The default is the Powershell icon.

.PARAMETER FormWidth
Use this parameter to set an alternate Menu Form width, 
the height is calculated at run-time.

.PARAMETER ShowTestMenu
Switch to cause Test Menu Display. May be paired with the DynamicParam 'MenuItemCount'.
'MenuItemCount' (Alias: MIC) defaults to 20. It's the number of items to display in the Test Menu.

.EXAMPLE
Show-DynamicMenu.ps1 -MenuTitles <String[]>
This example will display a test menu with an item for each array element.

.EXAMPLE
Show-DynamicMenu.ps1 -ST -MIC 12
This example will display a test menu with 12 items.
#>
[CmdletBinding()]
Param(
	[Parameter()][Alias('M')][String[]]$MenuTitles,
	[Parameter()][Alias('T')][String]$Title = 'Powershell Dynamic Menu',
	[Parameter()][Alias('MT')][String]$MenuTitle = 'Main Menu',
	[Parameter()][Alias('I')][Drawing.Icon]$Icon,
	[Parameter()][Alias('FW')][Int]$FormWidth = 250,
	[Parameter()][Alias('ST')][Switch]$ShowTestMenu)
#region Begin dynamic parameter definition
DynamicParam{
# Set up the Run-Time Parameter Dictionary
$RuntimeParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::New()

if($ShowTestMenu.IsPresent){
	#region MenuItemCount 
	$DynamicParamName = 'MenuItemCount'
	$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::New()
	$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::New()
	$ParameterAttribute.Mandatory = $false
	$ParameterAttribute.Position = 6
	$AttributeCollection.Add($ParameterAttribute)
	$ParameterAlias = [System.Management.Automation.AliasAttribute]::New('MIC')
	$AttributeCollection.Add($ParameterAlias)
	$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::New($DynamicParamName, [UInt16], $AttributeCollection)
	$RuntimeParameter.Value = 20
	$RuntimeParameterDictionary.Add($DynamicParamName, $RuntimeParameter)
}
# When done building dynamic parameters, return
return $RuntimeParameterDictionary
}
#endregion Dynamic parameter definition
Begin {
	#region Add Types
	Add-Type -A System.Windows.Forms
	Add-Type -A System.Drawing
	Add-Type -A Microsoft.VisualBasic
	$VBI = [Microsoft.VisualBasic.Interaction]
	#endregion

	#region Module Import
	Import-Module -Name .\Exists.ps1 -Force
	Import-Module -Name .\WinFormsLibrary.ps1 -Force
	#endregion

	#region Add BlueflameDynamics.IconTools Class
	$DLLPath = '.\BlueflameDynamics.IconTools.dll'
	if((Test-Exists -Mode File -Location $DLLPath)){Add-Type -Path $DLLPath}
	#endregion

	#region Script Level Variables
	[Int16]$SelectedItem = -1 #Return Value
	[Int16]$MaxMenuItems = [Int16]::MaxValue
	[Int]$DefaultFormWidth = 250
	if($ShowTestMenu.IsPresent){$ItemCount = $RuntimeParameterDictionary.Item('MenuItemCount').Value}
	#endregion

	#region Utility Functions
	Function Set-MenuItem{
	Param(
		[Parameter(Mandatory)][String[]]$Labels,
		[Parameter(Mandatory)][Windows.Forms.ToolStripMenuItem[]]$MenuItems,
		[Parameter(Mandatory)][Drawing.Size]$ItemSize,
		[Parameter(Mandatory)][String]$ItemPrefix,
		[Parameter()][String[]]$HotKeys = @(),
		[Parameter()][Switch]$SetSizeOff,
		[Parameter()][Switch]$NoHotKeys)

	for($C = 0; $C -le $MenuItems.GetUpperBound(0); $C++){
		$MenuItems[$C].Name = $ItemPrefix + ($C+1)
		$MenuItems[$C].Text = $Labels[$C]
		if(!$SetSizeOff){$MenuItems[$C].Size = $ItemSize}
		if(!$NoHotKeys){
			$MenuItems[$C].ShortcutKeys =`
			Get-ShortcutKey -Mode $(if($Hotkeys.Length -eq 0) {$Labels[$C]}else{$HotKeys[$C]})
			}
		}
	}

	Function Limit-MenuItems{
	Param([Parameter(Mandatory)][Alias('C')][Int]$Count)
	if($Count -gt $MaxMenuItems){
		Throw [System.ArgumentOutOfRangeException]::New(
			'Count',$Count,('Menu Item Count Limit of {0} Exceeded' -f $MaxMenuItems))}
	}
	#endregion

	<#
	.NOTES
	Name:	 Show-DynamicMenu
	Author:  Randy Turner
	Version: 1.0
	Date:	 11/16/2022

	.SYNOPSIS
	This function is designed to provide a WinForm Menu whose items are dynamically defined.

	.DESCRIPTION
	This function is designed to provide a WinForm Menu whose items are dynamically defined.
	The return integer value is the index value of the selected item in the input array or
	-1 for the Form Close Button(X), or -2 for the predefined menu Exit item\Alt-F4.

	.PARAMETER MenuTitles
	This is a String Array of the title text to appear in the menu. 
	This array is limited to 32767 items, surely more than needed.

	.PARAMETER Title
	This parameter is a String to appear in the form Title Bar.

	.PARAMETER MenuTitle
	This parameter is a String to appear in the menu dropdown Title.

	.PARAMETER Icon
	This parameter is an alternate Drawing.Icon to appear on the form.
	The default is the Powershell icon.

	.PARAMETER FormWidth
	Use this parameter to set an alternate Menu Form width, 
	the height is calculated at run-time.

	.EXAMPLE
	Show-DynamicMenu -MenuTitles <String[]>
	This example will display a menu with an item
	for each array element.
	#>
	Function Show-DynamicMenu{
	Param(
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][Alias('ME')][String[]]$MenuTitles,
		[Parameter()][Alias('T')][String]$Title = 'Powershell Dynamic Menu',
		[Parameter()][Alias('MT')][String]$MenuTitle = 'Main Menu',
		[Parameter()][Alias('I')][Drawing.Icon]$Icon,
		[Parameter()][Alias('FW')][Int]$FormWidth = $DefaultFormWidth)

	#region Function Level Variables
	[Int16]$Script:SelectedItem = -1 #Reset
	$Menu = @{
		ExitIndex = -1
		FormMinHeight = 60
		MaxItems = $Script:MaxMenuItems
		Multiplier = 0
		MaxMultiplier = 24
		ItemsSize = 0
		Text = @() 
		Items = @()
		SubItems = @()
	}
	#endregion

	#region Parameter Validation
	if($Null -eq $Icon){$Icon = [BlueflameDynamics.IconTools]::ExtractIcon($Env:SystemRoot+'\System32\ImageRes.dll',311,32)}
	Limit-MenuItems -Count $MenuTitles.Count
	#endregion

	#region Main Form
	$DMF = [Windows.Forms.Form]::New()
	$DMF.Text = $Title
	$DMF.Icon = $Icon
	$DMF.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
	$DMF.MaximizeBox = `
	$DMF.MinimizeBox = `
	$DMF.ShowInTaskbar = $False
	$DMF.Enabled = `
	$DMF.TopMost = $True
	$DMF.Size = [Drawing.Size]::New($FormWidth,$Menu.FormMinHeight)
	$DMF.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
	#endregion

	#region Main Menu
	$Menu.Text = @($MenuTitle)
	$Menu.Items = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count $Menu.Text.Count
	$Menu.SubItems = New-ObjectArray -TypeName Windows.Forms.ToolStripMenuItem -Count $MenuTitles.Count
	$Menu.ItemSize = [Drawing.Size]::New(218,22)
	$MainMenu = [Windows.Forms.MenuStrip]::New()
	$MainMenu.Name = 'DMF_MainMenu'
	$MainMenu.Parent = $DMF
	$MainMenu.Dock = [Windows.Forms.DockStyle]::Top
	$MainMenu.Size = [Drawing.Size]::New(220,30)
	$MainMenu.Items.AddRange($Menu.Items)
	Set-MenuItem $Menu.Text $Menu.Items $Menu.ItemSize 'MenuItem' -NoHotKeys 
	#endregion

	#region Main Menu SubItems
	$ItemClick = {
		$Script:SelectedItem = $Menu.SubItems.IndexOf($This)
		$DMF.Close()}
	$ExitClick = {
		$Script:SelectedItem = -2
		$DMF.Close()}
	Set-MenuItem $MenuTitles $Menu.SubItems $Menu.ItemSize 'MenuSubItem' -NoHotKeys
	$Menu.Items[0].DropDownItems.AddRange($Menu.SubItems)
	ForEach($Item in $Menu.SubItems){$Item.Add_Click($ItemClick)}
	#region Append Menu Exit & Separator
	$Menu.Multiplier = $VBI::IIf($Menu.Items[0].DropDownItems.Count -le $Menu.MaxMultiplier,$Menu.Items[0].DropDownItems.Count-1,$Menu.MaxMultiplier)
	[Void]$Menu.Items[0].DropDownItems.Add([Windows.Forms.ToolStripSeparator]::New())
	[Void]$Menu.Items[0].DropDownItems.Add([Windows.Forms.ToolStripMenuItem]::New())
	$Menu.ExitIndex = $Menu.Items[0].DropDownItems.Count - 1
	$Menu.Items[0].DropDownItems[$Menu.ExitIndex].Text = 'Exit'
	$Menu.Items[0].DropDownItems[$Menu.ExitIndex].Name = 'ExitItem'
	$Menu.Items[0].DropDownItems[$Menu.ExitIndex].ShortcutKeys = [Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::F4
	$Menu.Items[0].DropDownItems[$Menu.ExitIndex].Add_Click($ExitClick)
	$DMF.Add_Shown({$Menu.Items[0].ShowDropDown()})
	#endregion
	#region Limit Form & DropDown Height 
	$DMF.Size = [Drawing.Size]::New($FormWidth,$Menu.FormMinHeight+($Menu.ItemSize.Height*$Menu.Multiplier))
	$Menu.Items[0].DropDown.AutoSize = $False
	if($Menu.Multiplier -ge $Menu.MaxMultiplier){$Menu.Items[0].DropDown.Height = $DMF.Height}
	#endregion
	#endregion

	[Void]$DMF.ShowDialog()
	return $SelectedItem
	}

	<#
	.NOTES
	Name:	 Test-DynamicMenu
	Author:  Randy Turner
	Version: 1.0
	Date:	 11/16/2022

	.SYNOPSIS
	This function is designed to display a test\demo menu.

	.DESCRIPTION
	This function is designed to display a test\demo menu.
	When this script is imported as a module, use this
	function to display a test menu.
	#>
	Function Test-DynamicMenu{
	Param(
		[Parameter()][Alias('MI')][UInt16]$Items = 25,
		[Parameter()][Alias('T')][String]$Title = 'Dynamic Menu Test',
		[Parameter()][Alias('MT')][String]$MenuTitle = 'Test Menu',
		[Parameter()][Alias('I')][Drawing.Icon]$Icon,
		[Parameter()][Alias('FW')][Int]$FormWidth = $DefaultFormWidth)

	Limit-MenuItems -Count $Items
	$Mask = ('Test Menu Item: {0:D#}').Replace('#',$Items.ToString().Length)
	$TestItems = @(1..$Items|ForEach-Object{$Mask -f $_})
	$RV = Show-DynamicMenu -ME $TestItems -T $Title -MT $MenuTitle -I $Icon -FW $FormWidth
	Switch($RV){
		-1 {'Exited with Close Button'}
		-2 {'Exited by Menu Exit Item'}
		Default {'Selection: {0}' -f $TestItems[$RV]}
		}
	}
}
Process {
	# Called as Cmdlet?
	if($Null -ne $MenuTitles -and !$ShowTestMenu.IsPresent){
		Show-DynamicMenu -ME $MenuTitles -T $Title -MT $MenuTitle -I $Icon -FW $FormWidth}
	if($ShowTestMenu.IsPresent){Test-DynamicMenu -Items $ItemCount}
}