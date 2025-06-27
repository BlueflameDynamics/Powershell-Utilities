<#
.NOTES
Name:	Exists.ps1
Author:  Randy Turner
Version: 2.0a
Date:	09/24/2021
Revision History:
2.0a - 09/24/2021 - Updated for Powershell Core. Exists Deprecated in favor of Test-Exists.
1.0a - 06/15/2018 - Added Test-Exists function for Standard Naming Convention

.SYNOPSIS
Provides a wrapper for fumctions used to test for the existance of a file or directory

.PARAMETER Mode
Required mode of operation FILE\DIRECTORY

.PARAMETER Location
File\Directory to validate.

.EXAMPLE
Exists -Mode File -Location "c:\Video\PF_Save_Summer.mp4"
This example returns True if the file exists.

.EXAMPLE
Test-Exists -Mode Directory -Location "c:\Video\"
This example returns True if the directory exists.
#>
function Test-Exists{
	Param(
		[Parameter(Mandatory)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('Directory','File')]
			[String]$Mode,
		[Parameter(Mandatory)][String]$Location)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$ValidModes = $MyParam['Mode'].Attributes.ValidValues

	if($Location.StartsWith('.')){$Location = Resolve-CurrentLocation -Path $Location}

	Switch([Array]::IndexOf($ValidModes,$Mode))
		{
		0 {[IO.Directory]::Exists($Location)}
		1 {[IO.File]::Exists($Location)}
		}
}
function Resolve-CurrentLocation{
	param([Parameter(Mandatory)][Alias('P')][String]$Path)
	return "{0}\{1}" -f (Get-Location),(Split-Path -Path $($Path) -Leaf)
}

#Assign an Alias for Deprecated function 'Exists'
Set-Alias -Name Exists -Value Test-Exists