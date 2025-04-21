<#
.NOTES
	File Name:	Get-ColorInfo.ps1
	Version:	1.3 - 12/20/2022
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	06/10/2019

.SYNOPSIS
	Contains a function for Exporting Windows Color Set Information. 

.DESCRIPTION
	This script contains functions to Convert the KnownColors Enum into 3 ArrayLists with the names &
	ColorIndex (1-Based) values for The 174 KnownColors, Further splitting the KnownColors into the
	Subsets for  System colors(33) and WebColors(141). Functions are also provided for generating
	ArrayLists of the HSL and Argb values for the 3 color sets, and a Arraylist that contains
	PSObjects detailiing all the information relating to a color set. The Names in these ArrayLists
	may be used in a ForEach-Object Loop to get the corresponding [System.Drawing.Color] by use of
	commands like: $Color = [System.Drawing.Color]::FromName($_.Name). Using this ability a final
	option allows you to get an ArrayList of PSObjects that correspond to Powershell's native 
	representation of a [System.Drawing.Color].

.PARAMETER ColorSet - Alias: CS
	Select the Output Color Set: Known, System, or Web.

.PARAMETER OutputType - Alias: OT
	Sets the data to be output: 
		ColorSet:		ArrayList of Native Powershell Color Objects
		ColorSetInfo:	ArrayList of Color Names & ColorIndex values
		ColorInfoTable: ArrayList of All Color Property Values
		ArgbInHex:		ArrayList of Color Alpha,Red,Green,Blue channels expressed as a single Hex value
		HueList:		ArrayList of Color Names & HSL Hue values
		SaturationList:	ArrayList of Color Names & HSL Saturation values
		BrightnessList:	ArrayList of Color Names & HSL Lightness values

.OUTPUTS
	This script outputs an ArrayList of PSObjects for the requested OutputType.

.EXAMPLE
	.\Get-ColorInfo.ps1
	Default output of Windows KnownColors, Names & Color Index Values in Color Index order.

.EXAMPLE
	.\Get-ColorInfo.ps1 -ColorSet Web -OutputType ArgbInHex
	Outputs the Web Color Set Names & Alpha, Red, Green, Blues values of each as a hexidecimal value in Color Index order.

.EXAMPLE
	.\Get-ColorInfo -CS Web -OT ColorInfoTable|Sort-Object -Property Brightness -Descending|Format-Table
	Output a formatted Table of Color Info sorted by Brightness (HSL Liightnes) in decending order.
#>

[CmdletBinding()]
	param(
		[Parameter()][Alias('CS')]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('Known','System','Web')]
			[String]$ColorSet = 'Known',
		[Parameter()][Alias('OT')]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('ColorSet','ColorSetInfo','ColorInfoTable',
				'ArgbInHex','HueList','BrightnessList','SaturationList')]
			[String]$OutputType='ColorInfoTable')

#region Utility Functions
function New-ColorItem{
param(
		[Parameter(Mandatory)][Alias('N')][String]$Name,
		[Parameter(Mandatory)][Alias('V')][Object]$Value)
	 [PSCustomObject][Ordered] @{Name = $Name;Value = $Value}
}

function Get-ColorArgbValuesInHex{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Idx = 0
$Arl|ForEach-Object{
	$Color = [System.Drawing.Color]::FromName($_.Name)
	$ArlOut.Add((New-ColorItem -Name $Arl[$Idx].Name -Value $("{0,8:X8}" -f $Color.ToArgb())))|Out-Null
	$Idx++
	}
Return ,$ArlOut
}

function Get-ColorBrightnessValues{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Idx = 0
$Arl|ForEach-Object{
	$Color = [System.Drawing.Color]::FromName($_.Name)
	$ArlOut.Add((New-ColorItem -Name $Arl[$Idx].Name -Value $Color.GetBrightness()))|Out-Null
	$Idx++
	}
Return ,$ArlOut
}

function Get-ColorHueValues{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Idx = 0
$Arl|ForEach-Object{
	$Color = [System.Drawing.Color]::FromName($_.Name)
	$ArlOut.Add((New-ColorItem -Name $Arl[$Idx].Name -Value $Color.GetHue()))|Out-Null
	$Idx++
	}
Return ,$ArlOut
}

function Get-ColorSaturationValues{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Idx = 0
$Arl|ForEach-Object{
	$Color = [System.Drawing.Color]::FromName($_.Name)
	$ArlOut.Add((New-ColorItem -Name $Arl[$Idx].Name -Value $Color.GetSaturation()))|Out-Null
	$Idx++
	}
Return ,$ArlOut
}

<#
.NOTES
	File Name:	Get-ColorValueTable
	Version:	1.3 - 12/20/2022
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	06/10/2019

.SYNOPSIS
	Gets an ArrayList of Custom ColorInfo Objects

.DESCRIPTION
	Gets an ArrayList of Custom ColorInfo Objects based upon input ArrayList.

.PARAMETER Arl Alias: A
	Input ArrayList of Color Objects

.OUTPUTS
	ArrayList of Custom ColorInfo Objects.
	ArrayList ScriptMethod 'FindColorInfo([Drawing.Color]Color)' Search.
	Object ScriptMethod 'GetPSCustomObjectType()' returns 'BlueflameDynamics.ColorInfo'

.EXAMPLE
	Get-ColorValueTable -Arl $Arl
#>
function Get-ColorValueTable{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Arl|ForEach-Object{
	$Color = [Drawing.Color]::FromName($_.Name)
	$ColorObj =  [PSCustomObject][Ordered] @{
	PSTypeName = 'BlueflameDynamics.ColorInfo'
	Name = $_.Name
	Index = $_.Value
	A = $Color.A
	R = $Color.R
	G = $Color.G
	B = $Color.B
	ARGB = $("{0,8:X8}" -f $Color.ToArgb())
	Hue	= $("{0:000.0000000}" -f $Color.GetHue())
	Saturation = $("{0:0.0000000}" -f $Color.GetSaturation())
	Brightness = $("{0:0.0000000}" -f $Color.GetBrightness())
	IsSystemColor = $Color.IsSystemColor
	IsWebColor = ($_.Value -ge 27 -and $_.Value -le 167)}
	Add-Member -InputObject $ColorObj -MemberType ScriptMethod -Name GetPSCustomObjectType -Value {$This.PSObject.TypeNames[0]}
	$ArlOut.Add($ColorObj)
	} | Out-Null
Add-Member -InputObject $ArlOut -MemberType ScriptMethod -Name FindColorInfo -Value {
	param([Parameter(Mandatory)][Drawing.Color]$Color)	
	$RV = $Null
	ForEach($Entry In $This){
		If($("{0,8:X8}" -f $Color.ToArgb()) -eq $Entry.Argb){
	 		$RV = $Entry
			Return $RV}
	}
}
Return ,$ArlOut
}

function Get-PSColors{
param([Parameter(Mandatory)][Alias('A')][System.Collections.ArrayList]$Arl)
$ArlOut = [System.Collections.ArrayList]::New()
$Arl|ForEach-Object{
	$Color = [Drawing.Color]::FromName($_.Name)
	$ArlOut.Add($Color)
	} | Out-Null
Return ,$ArlOut
}
#endregion

#region Common Variables
$MyParam = ($MyInvocation.MyCommand).Parameters
$ColorSets = $MyParam['ColorSet'].Attributes.ValidValues
$OutputTypes = $MyParam['OutputType'].Attributes.ValidValues
$Arl = [System.Collections.ArrayList]::New()
$RV = [System.Collections.ArrayList]::New()
#endregion

#region Color Set ArrayLists
$ArlKnownColors = [System.Collections.ArrayList]::New()
$ArlSystemColors = [System.Collections.ArrayList]::New()
$ArlWebColors = [System.Collections.ArrayList]::New()
#endregion

#region ArrayList Initialization 
# Get KnownColors
foreach ($v in [Enum]::GetValues([System.Drawing.KnownColor])) {$ArlKnownColors.Add((New-ColorItem -N $v -V ([int]$v)))| Out-Null} 

#Derive System Color Set from KnownColors
$ArlSystemColors.AddRange($ArlKnownColors)
$ArlSystemColors.RemoveRange(26,141)  #Remove Web Colors

#Derive Web Color Set from KnownColors
$ArlWebColors.AddRange($ArlKnownColors)
$ArlWebColors.RemoveRange(167, 7) #Remove Trailing System Colors
$ArlWebColors.RemoveRange(0, 26)  #Remove Leading System Colors
#endregion

#region Main Process
Switch([Array]::IndexOf($ColorSets,$ColorSet)){
	0 {$Arl = $ArlKnownColors}
	1 {$Arl = $ArlSystemColors}
	2 {$Arl = $ArlWebColors}
	}

Switch([Array]::IndexOf($OutputTypes,$OutputType)){
	0 {$RV = Get-PSColors -Arl $Arl}
	1 {$RV = $Arl}
	2 {$RV = Get-ColorValueTable -Arl $Arl}
	3 {$RV = Get-ColorArgbValuesInHex -Arl $Arl}
	4 {$RV = Get-ColorHueValues -Arl $Arl}
	5 {$RV = Get-ColorBrightnessValues -Arl $Arl}
	6 {$RV = Get-ColorSaturationValues -Arl $Arl}
	}
#endregion

#Pass back Return Value ArrayList
Return ,$RV