<#
.NOTES
-------------------------------------
Name:	MetadataIndexLib.ps1
Version: 1.1 - 11/01/2021
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
Revision History:
Version: 1.1 - 11/01/2021 ~ Simplified Code & Added Documentation.
Version: 1.0 - 11/01/2007 ~ Initial Release

.SYNOPSIS
This script contains simple functions for accessing Extended File Properties. 
----------------------------------------------------------------------------------------

.DESCRIPTION
This library of functions makes use of a Shell.Application COM Object to
access Extended File Properties and build a WinOS Version independent MetaData Index
using the Shell32 Folder.GetDetailsOf() method. Two functions:
Get-MetadataIndex & Get-IndexByMetadataName are useful for building more complex
functions for retrieving the index value of a given named property or named group of 
properties. Such functions may be found in other scripts which make use of this library.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
#>

#region Script Level Variables
# $ShellObj is a custom object of Shell.Application & control values
$ShellObj = [PSCustomObject][Ordered]@{
	Path   = ''
	Folder = ''
	File   = ''
	Shell  = ''
	ShellFolder = ''
	ShellFile   = ''
	MaxPropertyIndex = 500 #Default Value
	Initialized = $False
}
#endregion

<#
.NOTES
Name:	 Init-ShellObj Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function initializes a custom object of Shell.Application & control values.

.PARAMETER TargetPath Alias: P
Optional, File path to base ShellObj upon. Must be a file, if a directory the index will
be for the directory's parent directory.

.PARAMETER MPI Alias: M
Optional, An Integer value that defines the MaxPropertyIndex, this varys with different values
with Windows versions 7 thru 10 has a range of 0-320 for the sake of future expansion the
MaxPropertyIndex defaults to 500.
#>
function Init-ShellObj
{
	param(
		[Parameter()][Alias('P')][String]$TargetPath='.',
		[Parameter()][Alias('M')][Int]$MPI=$ShellObj.MaxPropertyIndex)
	#Create Windows Shell Object
	$ShellObj.Path = $TargetPath
	$ShellObj.Shell = New-Object -COMObject Shell.Application
	$ShellObj.Folder = Split-Path -Path $ShellObj.Path
	$ShellObj.File = Split-Path -Path $ShellObj.Path -Leaf
	$ShellObj.ShellFolder = $ShellObj.Shell.Namespace($ShellObj.Folder)
	$ShellObj.ShellFile = $ShellObj.ShellFolder.ParseName($ShellObj.File)
	$ShellObj.MaxPropertyIndex = $MPI
	$ShellObj.Initialized = $True
}

<#
.NOTES
Name:	 Get-MetadataIndex Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function gets an array of MetadataIndexItems with a property Name & Index.

.PARAMETER TargetPath Alias: P
Optional, File path to base ShellObj upon. Must be a file, if a directory the index will
be for the directory's parent directory.

.PARAMETER Unfiltered Alias: U
Optional, A Switch that if present causes an Unfiltered MetadataIndex to be returned,
the default is to filter out unnamed index values as in most cases these are reserved
for future use. Some like 296 will return a value but are unnamed the MetadataIndex values
may vary with each directory and PC to PC. Some applications may extend the index with
custom properties of their own.
#>
function Get-MetadataIndex
{
	param(
		[Parameter(Mandatory)][Alias('P')][String]$TargetPath='',
		[Parameter()][Alias('U')][Switch]$Unfiltered)
	#To get a list of index numbers and their named properties
	Init-ShellObj $TargetPath
	$MetadataIndex = 0..$ShellObj.MaxPropertyIndex | Foreach-Object {New-MetadataIndexItem `
		-Index $_ `
		-Name  $ShellObj.ShellFolder.GetDetailsOf($null, $_)}
	if(!$Unfiltered){$MetadataIndex = $MetadataIndex|Where-Object -Property Name -ne -Value ''}
	$MetadataIndex
}

#Assign an Alias
Set-Alias -Name Show-MetadataIndex -Value Get-MetadataIndex

<#
.NOTES
Name:	 New-MetadataIndexItem Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function gets a custom MetadataIndexItem object with a property Name & Index.

.PARAMETER Name Alias: N
Name to include in output object.

.PARAMETER Index Alias: I
Integer to include as index value in output object.
#>
function New-MetadataIndexItem
{
	param(
		[Parameter()][Alias('N')][String]$Name='',
		[Parameter()][Alias('I')][Int]$Index=0)
	[PSCustomObject][Ordered]@{Index = $Index;Name = $Name}
}

<#
.NOTES
Name:	 Get-IndexByMetadataName Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function gets the Index value associated with property name.

.PARAMETER MetaDataIndex Alias: MDI
Required, An array of MetadataIndexItems to search

.PARAMETER SearchValue) Alias: SV
Required, Property Name to seek the index value.
#>
function Get-IndexByMetadataName
{
	param(
		[Parameter(Mandatory)][Alias('MDI')][Array]$MetaDataIndex,
		[Parameter(Mandatory)][Alias('SV')][String]$SearchValue)
	$MetaDataIndex.Index[$MetaDataIndex.Name.IndexOf($SearchValue)]
}

<#
.NOTES
Name:	 Get-ExtendedFileProperties Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function gets an array of custom PropertyItem objects with the file properties of the input file.

.PARAMETER Path Alias: P
Required, A fully qualified filename of the input file.

.PARAMETER MPI Alias: M
Optional, Maximum property index value to return.
Defaults to current $ShellObj.MaxPropertyIndex value
#>
function Get-ExtendedFileProperties
{
	param(
		[Parameter()][Alias('P')][String]$Path,
		[Parameter()][Alias('M')][Int]$MPI=$ShellObj.MaxPropertyIndex)
	$ShellObj.MaxPropertyIndex = $MPI
	Init-ShellObj $Path
	0..$MPI | Where-Object {$ShellObj.ShellFolder.GetDetailsOf($ShellObj.ShellFile, $_)} | 
		Foreach-Object{
			New-PropertyItem `
				-I $_ `
				-D $ShellObj.ShellFolder.GetDetailsOf($null, $_) `
				-V $ShellObj.ShellFolder.GetDetailsOf($ShellObj.ShellFile, $_)}
}

<#
.NOTES
Name:	 New-PropertyItem Function
Author:  Randy Turner
Version: 1.0
Date:	 11/01/2007

.SYNOPSIS
This function gets a custom PropertyItem object with 3 properties:
1. Description, the Property Name as it appears in the MetadataIndex.
2. Index, the index value of the property.
3. Value, A String containing the property value.

.PARAMETER Index Alias: I
MetadataIndex index value to include.

.PARAMETER Description Alias: D
Property Name to include.

.PARAMETER Value Alias: V
Property value ti include

.OUTPUT
A custom PropertyItem with the Index, Description (Property Name) & Value
of a single file property.
#>
function New-PropertyItem
{
	param(
		[Parameter()][Alias('I')][Int]$Index=0,
		[Parameter()][Alias('D')][String]$Description='',
		[Parameter()][Alias('V')][String]$Value='')
	[PSCustomObject][Ordered]@{Index = $Index;Description = $Description;Value=$Value}
}

<#
Import-Module .\MetadataIndexLib.ps1 -Force
Get-ExtendedFileProperties -P "\\MYCloud1\Public\Shared Videos\Movies\Avatar.m4v"|Out-GridView
Get-MetadataIndex -P "\\MYCloud1\Public\Shared Videos\Movies\"|Out-GridView
Show-MetadataIndex -P "\\MYCloud1\Public\Shared Videos\Movies\Avatar.m4v"|Out-GridView
#>
