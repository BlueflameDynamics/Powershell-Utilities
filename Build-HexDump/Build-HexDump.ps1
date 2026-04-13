<#
.NOTES
-------------------------------------
Name:        Build-HexDump.ps1
Version:     1.2 - 04/12/2026
Author:      Randy E. Turner
Email:       turner.randy21@yahoo.com
-------------------------------------
--------------------------------------------------------------------------------------------
Revision History:
v1.2 - 04/12/2026 - Improved VSCode support (deterministic open, safe delete, async handling)
v1.1 - 04/08/2026 - Added VSCode Support
v1.0 - 01/08/2022 - Original Release.
--------------------------------------------------------------------------------------------

.SYNOPSIS
Produces a Hex File Dump for any input file.

.DESCRIPTION
Uses Format-Hex to generate a hex dump and supports output to File, Host, or Printer.

When -OutputTo File (default), two dynamic switches are available:
-OutputToTemp (-T) and -OpenInEditor (-Edt).

When -OutputToTemp is used, output is created in $Env:Temp, optionally opened in an editor,
and then deleted safely (ISE or VSCode aware).

When -OpenInEditor is used, the temp file is opened in ISE or VSCode instead of Notepad.

When output is directed to a Printer, a GridView is shown to select the printer.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------
#>
[CmdletBinding()]
param(
	[Parameter(
		Position = 0,
		HelpMessage = 'The path of the input file',
		ValueFromPipeline = $True)]
		[String]$FullName,
	[Parameter(
		Position = 1,
		HelpMessage = 'Output Target')]
		[Alias('O')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('File','Host','Printer')]
		[String]$OutputTo = 'File'
)
DynamicParam {
	$RuntimeParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

	if($OutputTo -eq 'File'){
		#region OutputToTemp
		$DynamicParameterName = 'OutputToTemp'
		$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
		$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::new()
		$ParameterAlias = [System.Management.Automation.AliasAttribute]::new('T')
		$AttributeCollection.Add($ParameterAlias)
		$ParameterAttribute.Mandatory = $False
		$ParameterAttribute.HelpMessage = 'Output file to Temp Directory?'
		$AttributeCollection.Add($ParameterAttribute)
		$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::new(
			$DynamicParameterName,([switch]),$AttributeCollection)
		$RuntimeParameterDictionary.Add($DynamicParameterName, $RuntimeParameter)
		#endregion
		#region OpenInEditor
		$DynamicParameterName = 'OpenInEditor'
		$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
		$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::new()
		$ParameterAlias = [System.Management.Automation.AliasAttribute]::new('Edt')
		$AttributeCollection.Add($ParameterAlias)
		$ParameterAttribute.Mandatory = $False
		$ParameterAttribute.HelpMessage = 'Open Temp File in Text Editor?'
		$AttributeCollection.Add($ParameterAttribute)
		$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::new(
			$DynamicParameterName,([switch]),$AttributeCollection)
		$RuntimeParameterDictionary.Add($DynamicParameterName, $RuntimeParameter)
		#endregion
	}
	return $RuntimeParameterDictionary
}
begin{
	#region Host Detection
	$HostID = @{
		IsISE    =  $null -ne $psISE
		IsVSCode = ($null -ne $psEditor -and $null -ne $psEditor.GetEditorContext())
	}
	$CurrentEditorFile = @{
		ISE    = if($HostID.IsISE)   {$psISE.CurrentPowerShellTab.Files.SelectedFile};
		VSCode = if($HostID.IsVSCode){$psEditor.GetEditorContext().CurrentFile}
	}
	#endregion
	#region Script Variables
	$HexDump = $null
	$MyParam = ($MyInvocation.MyCommand).Parameters
	$Outputs = $MyParam['OutputTo'].Attributes.ValidValues
	$Tmp     = $RuntimeParameterDictionary.Item('OutputToTemp').IsSet
	$OpenEd  = $RuntimeParameterDictionary.Item('OpenInEditor').IsSet
	$TextEditor = 'Notepad.exe'
	$DumpExt = '.HexDump'
	$EOFMark = "{0}{1}[ End-of-File ]{1}" -f ("`t"*4),('-'*10)
	#endregion

	function Resolve-CurrentLocation {
		param([Parameter(Mandatory)][Alias('P')][String]$Path)
		return "{0}\{1}" -f (Get-Location),(Split-Path -Path $Path -Leaf)
	}

	Clear-Host
}
process{
	#region Validate Input
	if ([string]::IsNullOrWhiteSpace($FullName)){
		if ($HostID.IsISE){
			$FullName = $CurrentEditorFile.ISE.FullPath
			if(!$CurrentEditorFile.ISE.IsSaved){$CurrentEditorFile.ISE.Save()}
		}
		elseif($HostID.IsVSCode){
			$FullName = $CurrentEditorFile.VSCode.Path
		}
	}

	if($FullName.StartsWith('.')){$FullName = Resolve-CurrentLocation -Path $FullName}

	try {[Void](Test-Path -Path $FullName)}
	catch {
		throw $_.Exception.Message
		exit
	}

	$FileFI = [IO.FileInfo]::new($FullName)
	if (!$FileFI.Exists) {
		throw "Input File: $($FileFI.Name) Not Found!"
		exit
	}
	#endregion

	#region Output File Naming
	if($Tmp){$FileOut = Join-Path -Path $Env:Temp -ChildPath ($FileFI.Name + $DumpExt)}
	else {$FileOut = $FileFI.FullName + $DumpExt}
	#endregion

	#region Output Control
	'Working ...'
	$HexOut = Format-Hex -Path $FullName
	if(!$MyInvocation.ExpectingInput){Clear-Host}

	switch([Array]::IndexOf($Outputs,$OutputTo)){
		0 { #File
			($HexOut += $EOFMark.Substring(2)) | Out-File -FilePath $FileOut
			if($Tmp -and $OpenEd){
				if($HostID.IsISE){
					$psISE.CurrentPowerShellTab.Files.Add($FileOut)
					$HexDump = $psISE.CurrentPowerShellTab.Files[
						$psISE.CurrentPowerShellTab.Files.Count - 1]
				}
				elseif($HostID.IsVSCode){
					# Open file
					$psEditor.Workspace.OpenFile($FileOut)

					# Wait for VSCode to actually load it
					$WaitCount = 0
					while($psEditor.GetEditorContext().CurrentFile.Path -ne $FileOut -and $WaitCount -lt 20){
						Start-Sleep -Milliseconds 150
						$WaitCount++
					}
				}
				else{& $TextEditor $FileOut}

				# Safe delete loop (bounded)
				$MaxWait = 20
				$Count = 0
				while ((Test-Path $FileOut) -and $Count -lt $MaxWait){
					try {Remove-Item -LiteralPath $FileOut -ErrorAction Stop}
					catch {
						Start-Sleep -Milliseconds 250
						$Count++
					}
				}
			}
		}
		1 { # Host
			($HexOut += $EOFMark) | Out-Host
		}
		2 { # Printer
			$SelectedPrinter = Get-Printer | Out-GridView -Title 'Select Printer' -PassThru
			if ($null -ne $SelectedPrinter) {
				Set-Clipboard -Value (Split-Path -Path $FileOut -Leaf)
				$HexOut += (' ' * 16 + $EOFMark.Substring(4))
				$HexOut | Out-Printer -Name $SelectedPrinter.Name
			}
		}
	}
	#endregion
}
end{
	if($HostID.IsISE -and $null -ne $HexDump){
		$psISE.CurrentPowerShellTab.Files.SetSelectedFile($HexDump)
	}
}

# .\Build-HexDump.ps1 -OutputTo File -OutputToTemp -OpenInEditor