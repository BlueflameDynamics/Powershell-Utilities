<#
.NOTES
-------------------------------------
Name:		Build-HexDump.ps1
Version:	1.0 - 01/08/2022
Author:		Randy E. Turner
Email:		turner.randy21@yahoo.com
-------------------------------------
--------------------------------------------------------------------------------------------
Revision History:
v1.0 - 01/08/2022 - Original Release.
--------------------------------------------------------------------------------------------

.SYNOPSIS
This script will produce a Hex File Dump for any input file.

.DESCRIPTION
This script uses the Powershell Format-Hex Cmdlet to convert input files to produce
a Hex File Dump. 

Parameters are provided to direct the resulting output to 1 of 3 primary targets: 
File, Host, or Printer. 

Two dynamic switches are available when -OutputTo is File (the default): 
-OutputToTemp (alias -T) & -OpenInEditor (alias -Edt).

When using the switch -OutputToTemp output is created in the $Env:Temp directory,
sent to an ISE Editor Tab or NotePad, before the temp file is deleted. 

The second dynamic switch -OpenInEditor may be used to override sending the temp file
to an ISE Editor Tab & send it to NotePad instead.

When output is directed to a Printer, a GridView is shown of all available printers. 
Select the desired printer and click <OK> or <Cancel> to proceed.

When -OutputTo Printer mode is used the output filename is copied to the Clipboard to
permit naming files generated because of the printer type i.e., PDF.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------

DYNAMIC PARAMETERS

-OutputToTemp [<SwitchParameter>]

	Alias:	T
	Required?					False
	Position?					named
	Default value				False
	Accept pipeline input?		false
	Accept wildcard characters? false

-OpenInEditor [<SwitchParameter>]

	Alias:	Edt
	Required?					False
	Position?					named
	Default value				False
	Accept pipeline input?		false
	Accept wildcard characters? false

.Parameter Path
Required, Path to input file. 

.Parameter OutputTo - Alias: O
Optional, Output Target\Type - Specfies the target device.

.INPUTS
A file to dump.

.OUTPUTS
A Hexidecimal File Dump. Named: <Original Path><Original Filename>.HexDump

.EXAMPLE
PS> .\Build-HexDump.ps1 -Path .\AudioPlayer.ps1 -OutputTo File -OutputToTemp
#>
[CmdletBinding()]
param(
	[Parameter(
		Position = 0,
		Mandatory=$False,
		HelpMessage = 'The path of the input file',
		ValueFromPipeline=$True)]
		[String]$Path,
	[Parameter(
		Position = 1,
		Mandatory=$False,
		HelpMessage = 'Output Target')]
		[Alias('O')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('File','Host','Printer')]
		[String]$OutputTo='File')
DynamicParam{
	# Set up the Run-Time Parameter Dictionary
	$RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
	#region Begin dynamic parameter definition
	if($OutputTo -eq 'File'){
		#region OutputToTemp 
		$DynamicParameterName = 'OutputToTemp'
		$AttributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object -TypeName System.Management.Automation.ParameterAttribute
		$ParameterAlias = New-Object -TypeName System.Management.Automation.AliasAttribute -ArgumentList 'T'
		$AttributeCollection.Add($ParameterAlias)
		$ParameterAttribute.Mandatory = $False
		#$ParameterAttribute.Position = 2
		$ParameterAttribute.HelpMessage = "Output file to Temp Directory?"
		$AttributeCollection.Add($ParameterAttribute)
		$RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList $DynamicParameterName,([switch]),$AttributeCollection
		$RuntimeParameterDictionary.Add($DynamicParameterName, $RuntimeParameter)
		#endregion
		#region OpenInEditor
		$DynamicParameterName = 'OpenInEditor'
		$AttributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
		$ParameterAttribute = New-Object -TypeName System.Management.Automation.ParameterAttribute
		$ParameterAlias = New-Object -TypeName System.Management.Automation.AliasAttribute -ArgumentList 'Edt'
		$AttributeCollection.Add($ParameterAlias)
		$ParameterAttribute.Mandatory = $False
		#$ParameterAttribute.Position = 3
		$ParameterAttribute.HelpMessage = "Open Temp File in Text Editor?"
		$AttributeCollection.Add($ParameterAttribute)
		$RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList $DynamicParameterName,([switch]),$AttributeCollection
		$RuntimeParameterDictionary.Add($DynamicParameterName, $RuntimeParameter)
		#endregion
	}
	# When done building dynamic parameters, return
	return $RuntimeParameterDictionary
}
begin{
	#region Script Level Variables
	$MyName = ($MyInvocation.MyCommand).Name
	$MyParam = ($MyInvocation.MyCommand).Parameters
	$Outputs = $MyParam['OutputTo'].Attributes.ValidValues
	$PsIseHost = 'Windows PowerShell ISE Host'
	$PsIseCurTab = if($Host.Name -eq $PsIseHost){$psISE.CurrentPowerShellTab.Files.SelectedFile}
	$TextEditor = 'NotePad.exe' #Path to Text Editor
	$DumpExt = '.HexDump' #Output File Extention
	$EOFMark = "{0}{1}[ End-of-File ]{1}" -f ("`t"*4),('-'*10)
	$OTT = $RuntimeParameterDictionary.Item('OutputToTemp').IsSet
	#endregion
	Clear-Host
}
process{
	#region Validate Input
	try{[Void](Test-Path -Path $Path)}
	catch{
		throw $_.Exception.Message
		exit}
	$FileFI = [IO.FileInfo]::New($Path)
	if(!$FileFI.Exists){
		throw ("Input File: {0} Not Found!" -f $FileFI.Name)
		exit}
	#endregion

	#region Output File Naming
	if($OTT){
		$FileOut = Join-Path -Path $Env:Temp -ChildPath ($FileFI.Name + $DumpExt)}
	else{
		$FileOut = ($FileFI.FullName + $DumpExt)}
	#endregion

	#region Output Control
	'Working ...'
	$HexOut = Format-Hex -Path $Path
	if(!$MyInvocation.ExpectingInput){Clear-Host}
	switch([Array]::IndexOf($Outputs,$OutputTo)){
		0	{#File
			if($OTT){ 
				if($Host.Name -eq $PsIseHost -and !$RuntimeParameterDictionary.Item('OpenInEditor').IsSet){
					($HexOut += $EOFMark)|Out-File -FilePath $FileOut
					$psISE.CurrentPowerShellTab.Files.Add($FileOut)}
				else{
					#Adjust for TabWidth, Notepad (8) vs ISE (4)
					($HexOut += $EOFMark.Substring(2))|Out-File -FilePath $FileOut
					& $TextEditor $FileOut}
				#Do-While to force wait for large files
				do{ Start-Sleep -Seconds 5
					Remove-Item -LiteralPath $FileOut -ErrorAction SilentlyContinue}
				while(Test-Path -LiteralPath $FileOut)}
			else{($HexOut += $EOFMark.Substring(2))|Out-File -FilePath $FileOut}}
		1	{#Host
			($HexOut += $EOFMark)|Out-Host}
		2	{#Printer
			$SelectedPrinter = Get-Printer|Out-GridView -Title 'Select Printer' -PassThru
			if($Null -ne $SelectedPrinter){
				Set-Clipboard -Value (Split-Path -Path $FileOut -Leaf)
				#Replace Tabs with Spaces
				$HexOut += (' '*16+$EOFMark.Substring(4))
				$HexOut|Out-Printer -Name $SelectedPrinter.Name
				}}
	}
	#endregion
}
end{if($Null -ne $PsIseCurTab){$psISE.CurrentPowerShellTab.Files.SetSelectedFile($PsIseCurTab)}
	'Dump Complete'}