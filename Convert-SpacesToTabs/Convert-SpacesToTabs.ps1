<#
.NOTES
-------------------------------------
Name:	Convert-SpacesToTabs.ps1
Version: 2.0 - 03/27/2021
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
--------------------------------------------------------------------------------------------
Revision History:
V2.0 - 03/27/2021 Added ability to accept piped input
V1.0 - 07/04/2020 Inital Release
--------------------------------------------------------------------------------------------

.SYNOPSIS
Converts leading spaces to tabs. If FileOut is ommitted, Creates a working output of FileIn 
with a ".tvx" extension and renames it to FileIn on completion.

.DESCRIPTION
Converts leading spaces to tabs, saving disk space & reducing the number of clusters required.
By default it also converts Unicode UTF-16 encoded files to UTF-8 No BOM. Windows encodes files
as UTF-16 for multi-language support, but UTF-8 is more efficient for English language files.
An option is provided for encoding in UTF-16.
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 

.Parameter FullName - Alias: In
Required, Input File path, May come from Pipeline.

.Parameter FileOut - Alias: Out
Optional Output file path, if omitted defaults to FileIn with the extension ".tvx". 
Ignored if input is via the pipeline and the AutoNaming option is Forced.

.Parameter Encoding - Alias: ES
Sets output file Encoding scheme of Default {utf8NoBOM} or Unicode (UTF-16)

.Parameter SpacesPerTab - Alias: SPT
Optional, Number of spaces to be replaced by a single Tab, default: 4

.Parameter ShowStats - Alias: R
Optional, Switch causes conversion stats to be output

.EXAMPLE
PS> .\Convert-SpacesToTabs.ps1 -FullName .\AudioPlayer.ps1 -ShowStats
Creates an .\AudioPlayer.ps1 file in UTF-8 where leading spaces have been replaced by tabs in a 4:1 ratio.
This can significantly reduce the file size as seen below:

File		SizeIn SizeOut SizeDiff
----		------ ------- --------
AudioPlayer  98496   49240   -49256

.EXAMPLE
PS> Get-ChildItems -Path F:\SamplePath\*.vb | .\Convert-SpacesToTabs.ps1 -ShowStats
Creates a new file in UTF-8 for each vb source file where leading spaces have been replaced by tabs in a 4:1 ratio.
The conversion from Unicode to UTF-8 cuts file size significantly, 16-bit\8-bit chars.
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory=$True,
		ValueFromPipeline=$True,
		ValueFromPipelineByPropertyName=$True)]
		[Alias("In")][String[]]$FullName,
	[Parameter(Mandatory = $False)][Alias('Out')][String]$FileOut = "",
	[Parameter(Mandatory = $False)][Alias('SPT')][Int32]$SpacesPerTab = 4,
	[Parameter(Mandatory = $False)][Alias('ES')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Default','Unicode')]
		[String]$Encoding = 'Default',
	[Parameter(Mandatory = $False)][Alias('R')][Switch]$ShowStats)

Begin
{
$Fi = @($null,$null)
}
Process
{
	Foreach($FileIn in $FullName)
	{
	if($PSCmdlet.MyInvocation.ExpectingInput){
		<#
		Input is via the pipeline
		Reset values each iteration
		#>
		$Fi = @($null,$null) 
		$FileOut = "" #Ignore FileOut, force AutoOut
		} 
	$AutoOut = ($FileOut.Length -eq 0) 
	$Fi[0] = Get-ChildItem -Path $FileIn

	If($Fi[0].Exists){
		If($AutoOut){
			$FileOut = -join ($Fi[0].DirectoryName,"\",[IO.Path]::GetFileNameWithoutExtension($Fi[0].Name),".tvx")
		}

		$Content = Get-Content -Path $FileIn|%{
		[Regex]::Replace($_,"( {$SpacesPerTab})","`t")}|
		Set-Content -Path $($FileOut) -Encoding $Encoding

		If($ShowStats.IsPresent){
			$Fi[1] = Get-ChildItem -Path $FileOut
			$Stat = New-Object -TypeName PSObject
			Add-Member -InputObject $Stat -MemberType NoteProperty -Name File -Value ([IO.Path]::GetFileNameWithoutExtension($Fi[1].Name))
			Add-Member -InputObject $Stat -MemberType NoteProperty -Name SizeIn -Value $Fi[0].Length
			Add-Member -InputObject $Stat -MemberType NoteProperty -Name SizeOut -Value $Fi[1].Length
			Add-Member -InputObject $Stat -MemberType NoteProperty -Name SizeDiff -Value ($Fi[1].Length - $Fi[0].Length)
			$Stat
		}

		If($AutoOut){
			$BakFile = -join ($FileIn,".bak")
			If(Test-Path -Path $BakFile){Remove-Item -LiteralPath $BakFile} 
			Rename-Item -Path $FileIn -NewName $BakFile
			Rename-Item -Path $FileOut -NewName $FileIn
		}
	} 
	else{
		Write-Error -Message $("Error: File - {0} Not Found!." -f $FileIn) -Category ResourceUnavailable
		}
	}
}
End{}