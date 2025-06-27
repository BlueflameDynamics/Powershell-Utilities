<#
.NOTES
-----------------------------------------------------------------
Name:	 Get-NetMediaInfo.ps1
Version:  4.2 - 09/16/2022
Author:   Randy E. Turner
Email:	turner.randy21@yahoo.com
Revision: This version supports full process automation and has
been tested using Windows Powershell v5.1 & Powershell Core V7.15+
v4.2 - Added a test to insure internet connectivity.
-----------------------------------------------------------------
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------

.SYNOPSIS
This Cmdlet will run my Export-MediaFileDetailsEx.ps1
collecting data for a fixed set of location(s) and
optionally import the results into a predefined 
MSAccess database, generate the requsted report(s), &
transfer a predefined set of files to a predefined location,
	
.DESCRIPTION
Network Media File inventory update Master. 

.PARAMETER Mode
Mode of Operation <Required> All, Audio, Image, or Video

.PARAMETER Publish Alias: P
Switch to enable importing gathered data into Access,
outputting the associated Access reports as PDF files,
and copying the I/O file to the NAS.

.PARAMETER Quiet Alias: Q
Switch to enable running Access Quietly.

.PARAMETER Reboot Alias: R
Switch to enable rebooting the host computer upon completion.

.EXAMPLE
To collect Video File Properties run Access Quietly & Reboot
Get-NetMediaInfo -Mode Video -P -Q -R

.EXAMPLE
To collect All Media File Data run Access Quietly & Reboot
Get-NetMediaInfo -Mode Video -Publish -Quiet -Reboot
#>

[CmdletBinding()]
param(
		[Parameter(Mandatory)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('All','Audio','Image','Video')]
			[String]$Mode,
		[Parameter()][Alias('P')][Switch]$Publish,
		[Parameter()][Alias('Q')][Switch]$Quiet,
		[Parameter()][Alias('R')][Switch]$Reboot)

#region Module Import
#Set-Location -Path <fully qualified path to NAS Inventory scripts>
Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\Test-InternetConnection.ps1 -Force
#endregion

#region Common Variables
$ScriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$MyParam = ($MyInvocation.MyCommand).Parameters
$Modes  = $MyParam['Mode'].Attributes.ValidValues
$ModeIdx = [Array]::IndexOf($Modes,$Mode)
$LogFile = [System.IO.Path]::GetFullPath(-join (".\",$ScriptName,"_Log.txt"))
$TargetServer = '\\MyCloud1\Public'
$EmailTxt = @(
	"$ScriptName Completed!",
	"$ScriptName of TPSI-NET Completed!",
	 $LogFile)
$Locations = @(
	"$TargetServer\Shared Music\",
	"$TargetServer\Shared Pictures\",
	"$TargetServer\Shared Videos\")
$ExportFile = @(
	'MediaCatalog.accdb',
	'MCAR001.pdf',
	'MCIR001.pdf',
	'MCVR001.pdf',
	'AudioReport.txt',
	'ImageReport.txt',
	'VideoReport.txt')
$CurrentLocation = Get-Location
$ExportFile = $ExportFile |ForEach-Object {[System.IO.Path]::GetFullPath(("{0}\{1}" -f $CurrentLocation,$_))}
$ExportTarget = "$TargetServer\Shared Applications\PowerShell\Scripts\"
#endregion

function Get-RegistryValue{
	param(
		[Parameter(Mandatory)][String]$Key,
		[Parameter(Mandatory)][String]$Value)
	(Get-ItemProperty -Path $Key).$Value
}

function Play-Sound{
	param(
		[Parameter(Mandatory)][Alias('F')][String]$SoundFile,
		[Parameter()][Alias('D')][int]$Delay=3)
$SoundLib = [System.IO.Path]::GetFullPath(("{0}{1}"-f (Get-Location),'\Audio\'))
$AudioFile = -join ($SoundLib,$SoundFile)
(New-Object -TypeName Media.SoundPlayer -ArgumentList "$AudioFile").Play()
Start-Sleep -Seconds $Delay
}

function Export-Files{
Play-Sound -SoundFile 'TransferringData.wav'
if(Test-Exists -Mode Directory $ExportTarget){
	For($C = 0;$C -le $ExportFile.GetUpperBound(0); $C++){
		if(Test-Exists -Mode File $ExportFile[$C]){
			Copy-Item -Path $ExportFile[$C] -Destination $ExportTarget -Force}
		}
	}
	Play-Sound -SoundFile 'DataTransferComplete.wav'
}

function Get-Audio{
.\Export-MediaFileDetailsEx -Mode Audio -Dir $Locations[0] -R -LC -LA -Log $EmailTxt[2]}
 
function Get-Image{
.\Export-MediaFileDetailsEx -Mode Image -Dir $Locations[1] -R -LC -LA -Log $EmailTxt[2]}

function Get-Video{
.\Export-MediaFileDetailsEx -Mode Video -Dir $Locations[2] -R -LC -LA -Log $EmailTxt[2]}

function Get-All{
$ShowStatus = {
	param([Int]$ModeIdx)
	if ($Host.Name -eq 'Windows PowerShell ISE Host'){
		Clear-Host
		"Step ($Modes[$ModeIdx]) $ModeIdx of 3 Complete!"
		#Flush Memory Variables
		Get-Variable -Exclude PWD,*Preference|Remove-Variable -EA 0
	}}
Get-Audio; Invoke-Command -Scriptblock $ShowStatus -ArgumentList 1
Get-Image; Invoke-Command -Scriptblock $ShowStatus -ArgumentList 2
Get-Video; Invoke-Command -Scriptblock $ShowStatus -ArgumentList 3
}

function Get-NetMediaInfo{
Import-Module -Name .\AES-Email.ps1 -Force
Import-Module -Name .\Suspend-PowerPlan.ps1 -Force
$TM = ''
$DbSw = ''
if(Test-Exists -Mode File $LogFile){Remove-Item -Path $LogFile}
Play-Sound -SoundFile 'Authoriz.wav' -Delay 5
Suspend-PowerPlan -System -Continuous
$TM = Switch($ModeIdx){
	0 {Measure-Command -Expression {Get-All}}
	1 {Measure-Command -Expression {Get-Audio}}
	2 {Measure-Command -Expression {Get-Image}}
	3 {Measure-Command -Expression {Get-Video}}
	}

if($Publish.IsPresent){
	$DbSw = Switch($ModeIdx){
	0 {'*'}
	1 {'A'}
	2 {'I'}
	3 {'V'}
	}
	if($Quiet.IsPresent){$DbSw = -join ($DbSw,'Q')}
	$KeyInfo = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msaccess.exe','(Default)')
	$AccPath = Get-RegistryValue -Key $KeyInfo[0] -Value $KeyInfo[1]
	$P=Start-Process -FilePath $AccPath -ArgumentList (([IO.Path]::GetFullPath($ExportFile[0])," /runtime /Cmd $DbSw")) -NoNewWindow -PassThru
	Do{<#Do Nothing#>} Until($P.HasExited -eq $True)
	Export-Files
	}

$TM|Out-File -FilePath $($EmailTxt[2]) -Append
if(Test-InternetConnection){
	if(Test-Exists -Mode File '.\AES.key'){
		Play-Sound -SoundFile 'TOS_Bosun_Whistle_1.wav'
		'Sending Confirmation Email, Please wait ...'| Out-Host
		Send-Email3 -Subject $EmailTxt[0] -Body $EmailTxt[1] -Att $EmailTxt[2]
		'Transmission Complete ...'| Out-Host
		Remove-Item  -Path $($EmailTxt[2])}
}
Suspend-PowerPlan #Reset
Play-Sound -SoundFile 'Analysis.wav' -Delay 5
if($Reboot.IsPresent){
	Play-Sound -SoundFile 'AutoShut.wav'
	Start-Sleep -Seconds 30
	Restart-Computer}
}

#Execute Main Function
Get-NetMediaInfo