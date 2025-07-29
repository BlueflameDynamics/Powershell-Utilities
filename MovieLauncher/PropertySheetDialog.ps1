<#
.NOTES
-------------------------------------
Name:	PropertySheetDialog.ps1
Version: 1.1 - 07/29/2025
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This script contains a function used to display a folder\file Property Sheet Dialog

.DESCRIPTION
This script contains a function used to display a folder\file Property Sheet Dialog
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
.Parameter TargetPath - Alias: P
The path of target folder\file.
#>

function Show-PropertySheetDialog
{
[CmdletBinding()]
param([Parameter(Mandatory)][Alias('P')][String]$TargetPath)

	#region Function Level Variables
	# $ShellObj is a custom object of Shell.Application & control values
	$ShellObj = [PSCustomObject][Ordered]@{
		Path   = ''
		Folder = ''
		File   = ''
		Shell  = $Null
		ShellFolder = $Null
		ShellFile   = $Null
	}
	$IsContainer = Test-Path -Path $TargetPath -PathType Container
	#endregion

	#region Create Windows Shell Object
	$ShellObj.Path = $TargetPath
	$ShellObj.Shell = New-Object -COMObject Shell.Application
	if ($IsContainer) {
		$ShellObj.Folder = $ShellObj.Path}
	else{
		$ShellObj.Folder = Split-Path -Path $ShellObj.Path}
	$ShellObj.File = Split-Path -Path $ShellObj.Path -Leaf
	$ShellObj.ShellFolder = $ShellObj.Shell.Namespace($ShellObj.Folder)
	$ShellObj.ShellFile = $ShellObj.ShellFolder.ParseName($ShellObj.File)
	#endregion

	#region Display approiate Property Sheet
	$Verb = "Properties"
	if ($IsContainer) {
		<# Folder #>	$ShellObj.ShellFolder.Self.InvokeVerb($Verb)
	} else {
		<# File #>		$ShellObj.ShellFile.InvokeVerb($Verb)
	}
	#endregion

	#region Clean-up Com objects
	if(!$IsContainer){$Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ShellObj.ShellFile)}
	$Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ShellObj.ShellFolder)
	$Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ShellObj.Shell)
	$ShellObj = $Null
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
	#endregion
}