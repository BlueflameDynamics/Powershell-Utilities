<#
.Notes
	-------------------------------------
	Name:	Set-OpticalDriveState
	Author:  Randy Turner
	Version: 2.0 - 07/20/2020
	Email:   turner.randy21@yahoo.com
	-------------------------------------
.SYNOPSIS
	This script helps in ejecting or closing a Optical Drive
	----------------------------------------------------------------------------------------
	Security Note: This is an unsigned script, Powershell security may require you run the
	Unblock-File cmdlet with the Fully qualified filename before you can run this script,
	assuming PowerShell security is set to RemoteSigned.
	---------------------------------------------------------------------------------------- 
.DESCRIPTION
	This script helps in ejecting or closing a Optical Drive
.PARAMETER Drive
	Drive Letter of target optical drive, If omitted targets first drive detected. An * selects all.
.PARAMETER Eject
	Ejects the target Drive
.PARAMETER Close
	Closes the target Drive
.PARAMETER Diagnostic
	Causes the display of a list of available optical drives, overrides other parameters
.EXAMPLE
	C:\PS>c:\Scripts\Set-OpticalDriveState -Drive D: -Eject
	Ejects the CD/DVD/BD Drive
.EXAMPLE
	C:\PS>c:\Scripts\Set-OpticalDriveState -Drive D: -Close
	Closes the CD/DVD/BD Drive
.EXAMPLE
	C:\PS>c:\Scripts\Set-OpticalDriveState -Diagnostic
	Displays profile of available optical drives
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory = $False,Position=0)][Alias('D')][String]$Drive='~',
	[Parameter(Mandatory = $False,Position=1)][Alias('E')][switch]$Eject,
	[Parameter(Mandatory = $False,Position=2)][Alias('C')][switch]$Close,
	[Parameter(Mandatory = $False,Position=3)][Alias('T')][switch]$Diagnostic
)

$DiscMaster = New-Object -ComObject IMAPI2.MsftDiscMaster2
$DiscRecorders = @()  #Available Recorders
$DiscRecorder = $Null #Selected Recorder
$Drive = $Drive.ToUpper()

Function Access-Drive{
	try{
		If($Eject.IsPresent){$DiscRecorder.EjectMedia()}
		ElseIf($Close.IsPresent){$DiscRecorder.CloseTray()}
	} 
	catch{Write-Error "Failed to operate the disc. Details : $_"}
}

#Format Drive Letter as: <letter>:\
Switch($Drive.Length){
	1 {$Drive += ":\"}
	2 {$Drive += "\" }
}

#Get available Optical Drives
For($Idx = 0;$Idx -lt $DiscMaster.Count;$Idx++){
	$UniqueId = $DiscMaster.Item($Idx)
	$DiscRecorder = New-Object -ComObject IMAPI2.MsftDiscRecorder2
	$DiscRecorder.InitializeDiscRecorder($UniqueId)
	$DiscRecorders += $DiscRecorder
}

#Default to 1st Drive
If($Drive.StartsWith('~')){
$Drive = $DiscRecorders[0].VolumePathNames[0]
}

If($Diagnostic.IsPresent){
	$DiscRecorders
	Exit
}

If(!$Drive.StartsWith('*')){
	$DiscRecorder = $Null #Reset for selection
	ForEach($Recorder In $DiscRecorders){
		if($Recorder.VolumePathNames[0] -eq $Drive){
		$DiscRecorder = $Recorder
		Break
		}
	}
	If($Null -ne $DiscRecorder){
	Access-Drive
	Exit
	}
	Else{Write-Error -Message "Drive: $Drive is not an optical drive" -Category InvalidArgument;Exit}
	}
ForEach($DiscRecorder In $DiscRecorders){
	Access-Drive
}
