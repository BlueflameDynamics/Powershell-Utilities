<#
.NOTES
Name:	Invoke-CopyFile.ps1
Author:  Randy Turner
Version: 1.0
Date:	11/04/2022

.SYNOPSIS
This script is designed to copy files and show a GUI Progress Window, when needed.
It may be called as a Cmdlet or Imported as a Module to allow direct access to it's functions.
.DESCRIPTION
Powershell Function Library Script

.PARAMETER Source
The source file path

.PARAMETER Target
The target file\directory path

.PARAMETER Overwrite
This switch will cause an existing file to be automatically overwritten.

.PARAMETER AsJob
This switch will cause the copy operation to run as a job.
#>
[CmdletBinding()]
Param(
		[Parameter()][String]$Source,
		[Parameter()][String]$Target,
		[Parameter()][Switch]$Overwrite,
		[Parameter()][Switch]$AsJob)

Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\GetSplitPathLib.ps1 -Force

<#
.NOTES
Function Name:	Invoke-CopyFile
Author:  Randy Turner
Version: 1.0
Date:	08/04/2007

.SYNOPSIS
This function is designed to copy a file using the
Microsoft.VisualBasic.FileIO.FileSystem.CopyFile() method
to display a GUI progress window and error message alerts.
This VB method may be called by any of the CLR languages.

.DESCRIPTION
This function is designed to copy a file and to provide design time overwrite control.

.PARAMETER Source
This parameter is the input file path.

.PARAMETER Target
This parameter is the output file or directory path.
When a directory only the Source file is copied to the target
using the Source FileName, Use of a fully qualified path will
permit renaming of the output file.

.PARAMETER Overwrite
Use of this switch will cause an existing Target file to be
overwritten without user interaction. 

.PARAMETER AsJob
Use of this switch will cause the copy operation to be run
asynchronously in a Fire & Forget fashion on a seperate thread
allowing the calling process to continue.

.EXAMPLE
Invoke-CopyFile -Source C:\Appdev\Powershell\Invoke-CopyFile.ps1 -Target \\MyServer\AppDev\Powershell -Overwrite
This example will copy the file to a new location overwriting any existing file.
 
.EXAMPLE
Invoke-CopyFile -Source C:\Appdev\Powershell\Invoke-CopyFile.ps1 -Target \\MyServer\AppDev\Powershell -Overwrite -AsJob
This example will copy the file to a new location overwriting any existing file asynchronously.
#>
function Invoke-CopyFile{
	Param(
		[Parameter(Mandatory)][String]$Source,
		[Parameter(Mandatory)][String]$Target,
		[Parameter()][Switch]$Overwrite,
		[Parameter()][Switch]$AsJob)

	$SF=$Source|Get-SplitPath
	$TF=$Target|Get-SplitPath
	if($TF.Extension.Length -eq 0 -and $TF.File.Length -ne 0)
		{$TF=Get-SplitPath ($Target += '\')}
	if($TF.File.Length -eq 0)
		{$TF=Get-SplitPath $(-Join($TF.Directory,'\',($TF.File=$SF.File)))}
	$CFB={param($SF,$TF)
		try	{
			Add-Type -A 'Microsoft.VisualBasic'
			[Microsoft.VisualBasic.FileIO.FileSystem]::CopyFile(
			$SF.FullPath,
			$TF.FullPath,
			[Microsoft.VisualBasic.FileIO.UIOption]::AllDialogs,
			[Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
			}
		catch {$_}
	}

	if($OverWrite.IsPresent -and (Test-Exists -Mode File $TF.FullPath) -and $TF.Extension.Length -gt 0)
			{Remove-Item -Path $TF.FullPath}

	if($AsJob)
		{Start-Job -ScriptBlock $CFB -ArgumentList $SF,$TF}
	else
		{& $CFB $SF $TF}
}

if($Source -ne ''){
	Invoke-CopyFile -Source $Source -Target $Target -Overwrite:$Overwrite -AsJob:$AsJob}