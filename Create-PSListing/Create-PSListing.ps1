<#
.NOTES
	File Name:	Create-PSListing.ps1
	Version:	2.0 - 03/25/2024
	Requires:	Powershell V5+
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	03/08/2016
	Based Upon: Copy-ToPrinter.ps1 by Karl Mitschke
	Revision History:
		V2.0 - 03/25/2024 - Merged in Create-TokenReport.ps1 V1.0 functions
		V1.9 - 02/01/2022 - Added Support for VS Code Editor like that of the ISE
		V1.8 - 01/24/2022 - Added Validate-ScriptType to simplify code
		V1.7 - 08/02/2020 - Added Do-While for Remove-Item, of Temp File for large file support.
		V1.6 - 09/17/2019 - Added Temp file clean-up when -OpenInWebBrowser & -OutputToTemp are paired.
		V1.5a - 10/19/2018
			- added support for .psm1 files
			- only allow supported types
			- support unsaved "UntitledN.ps1*" files in ISE
			- added option to output to a temporary file when used with -OpenInWebBrowser

.SYNOPSIS
	Converts the current script from the ISE\VSCode Editor or a file passed by the Path parameter
	from the console to an HTML file. Version 2+ will produce a Powershell Token report for
	ps1 & psm1 files similar to a Cross-Reference report for other programming languages.

.DESCRIPTION
	This script converts a PowerShell script to an HTML file in the same location as the original script or
	optionally in the directory defined by the $env:Temp value when output is displayed in the default web browser.
	This will allow printing a PowerShell listing from any application that supports printing HTML 
	eliminating the dependence of Copy-ToPrinter upon IE which is no longer supported by Microsoft.
	Version 2.0 adds the ability to alternatly\additionally produce a Token report for ps1 & psm1
	files similar to a Cross-Reference report for other programming languages. This combines the 
	functionality of eariler versions of Create-PSListing.ps1 & Create-TokenReport.ps1 into a single script.
	This script uses the Powershell parser to Tokenize an input script to produce a report.
	When producing a token-only report the output file name has '_Tokens' appended before the file extension.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------	

.PARAMETER Path
	Fully Qualyfied Script Name.
.PARAMETER Mode
	Mode of operation, may be Listing, Tokens, or Both.	
.PARAMETER FontSize - Alias: Fs
	 Listing font size to be used, between 10-24 points, default: 16.
.PARAMETER TabWidth - Alias: Tw
	Tab Expansion Width (2,4,6,8), default: 4.
.PARAMETER LineNoWidth - Alias: Lw
	Listing Line Number Width (3 thru 8), default: 4.	
.PARAMETER OutputTo - Alias: Out
	Optional, Token report Output Target\Type - Specfies the target device or Raw Output to default device.
.PARAMETER MaxContentLength - Alias: Max
	Optional, Specifies Token report max comment length to output, default: 50
.PARAMETER ShowComments - Alias: ICT
	Optional, Token report Switch causes Comment type tokens to be included.
.PARAMETER ShowStart - Alias: SS
	Optional, Token report Switch shows the 'Start' property for the "Script Buffer Offset" 
.PARAMETER OpenInWebBrowser - Alias: O
	Opens the output HTML file in the default web browser.
.PARAMETER OutputToTemp - Alias: T
	When used with -OpenInWebBrowser outputs to a temporary file.
.INPUTS
	Requires a file be open in the ISE\VS-Code, or a file path and name passed to the Path parameter.
.OUTPUTS
	This script outputs a formatted HTML file (<ScriptName>.htm) of the script with line numbers.
.EXAMPLE
	.\Create-PSListing.ps1 -Mode Both
	Creates listing of the current PowerShell ISE\VS-Code Editor tab with a token report.
.EXAMPLE
	.\Create-PSListing.ps1 -Path C:\scripts\Find-EmptyGroups.ps1
	Creates listing file: C:\scripts\Find-EmptyGroups.htm.
.EXAMPLE
	.\Create-PSListing.ps1 -Path C:\scripts\Find-EmptyGroups.ps1 -Mode Tokens -OutputTo HTML
	Creates token report file: C:\scripts\Find-EmptyGroups_Tokens.htm.
.EXAMPLE
	Get-ChildItem -Path C:\scripts\Find-DisabledMailbox.ps1|.\Create-PSListing.ps1 -Mode Both
	Creates an HTML file: C:\scripts\Find-DisabledMailbox.htm, with both a listing and token report.
#>

[CmdletBinding()]
param(
	[Parameter(
		Position = 0,
		ValueFromPipeline,
		HelpMessage = 'The path of the input file')][String]$Path,
	[Parameter(
		HelpMessage = 'Mode of Operation')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Listing','Tokens','Both')][String]$Mode = 'Listing',	
	[Parameter(HelpMessage = 'Listing HTML Font Size')][Alias('Fs')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(10,24)]
		[UInt32]$FontSize = 16,
	[Parameter(HelpMessage = 'Tab Expansion Width')][Alias('Tw')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet(2,4,6,8)]
		[UInt32]$TabWidth = 4,
	[Parameter(HelpMessage = 'Listing Line# Width')][Alias('Lw')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(2,8)]
		[UInt32]$LineNoWidth = 4,
	[Parameter(HelpMessage = 'Token Report Output Target, Raw Outputs Unfiltered Native Tokens')][Alias('Out')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('File','GridView','Host','Html','Raw')][String]$OutputTo = 'File',
	[Parameter(HelpMessage = 'Token Report Maximum Comment Length')][Alias('Max')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(10,100)][UInt32]$MaxContentLength = 50,
	[Parameter(HelpMessage = 'Token Report Include Comment Type?')][Alias('ICT')][Switch]$ShowComments,	
	[Parameter(HelpMessage = 'Token Report Include Start Property?')][Alias('SS')][Switch]$ShowStart,
	[Parameter(HelpMessage = 'Opens output file in default web browser')][Alias('O')]
		[switch]$OpenInWebBrowser,
	[Parameter(HelpMessage = 'When Used with -OpenInWebBrowser Outputs file to Temp Directtory')][Alias('T')]
		[switch]$OutputToTemp)

#Set-StrictMode -Version 5
Add-Type -Assembly System.Web	

#region Custom Enums
Enum ModeOfOperation{
	Listing
	Tokens
	Both
}
Enum TokenReportOutput{
	File
	GridView
	Host
	Html
	Raw
}
#endregion

#region Script Parameter Variables
$P = @(0,1,2)
$P[0] = ($MyInvocation.MyCommand).Parameters
$P[1] = $P[0]['Mode'].Attributes.ValidValues
$P[2] = $P[0]['OutputTo'].Attributes.ValidValues
$My = [PSCustomObject][Ordered]@{
	Name = ($MyInvocation.MyCommand).Name
	Params = $P[0]
	Modes = $P[1]
	Outputs = $P[2]
	ModeIdx = [Array]::IndexOf($P[1],$Mode)
	OutputIdx = [Array]::IndexOf($P[2],$OutputTo)}
# HTML required for combined report
if($My.ModeIdx -eq [ModeOfOperation]::Both){$OutputTo = 'Html'} 
RV -Name P
#endregion

#region Script Variables
$AppVersion = '2.0'
$psISEHost = 'Windows PowerShell ISE Host'
$VSCodeHost = 'Visual Studio Code Host'
$ReportName = 'Powershell Token Report'
$MatchStr = '\.psm?1$'
$RVStr = @('.psm*1','_Tokens.txt','\*$','','_Tokens.htm')
$ErrMsg = @(' only supports .ps1 and .psm1 files',' requires a .ps1 or .psm1 input file')
$LineNo = 1
$LineNoFormat = -join('{0:',('0' * $LineNoWidth),'}<BR/>')
$HtmlFont = -join('font-family:Consolas,Lucida Console; font-size:',$FontSize,'pt;')
$OutName = ''
$ScriptName = ''
$ScriptPath = ''
$Text = ''
$RawTokens = ''
$Tokens = ''
$TokenColors = @{
'Attribute' = 67
'Command' = 37
'CommandArgument' = 38
'CommandParameter' = 123
'Comment' = 53
'GroupEnd' = 35
'GroupStart' = 35
'Keyword' = 49
'LineContinuation' = 35
'LoopLabel' = 49
'Member' = 35
'NewLine' = 35
'Number' = 140
'Operator' = 68
'Position' = 35
'StatementSeparator' = 35
'String' = 59
'Type' = 158
'Unknown' = 35
'Variable' = 43
}
#endregion

#region Utility Functions
function Expand-Tabs{
param(
	[Parameter(
		Position = 0,
		ValueFromPipeline,
		Mandatory)][Alias('In')]
		[String]$InputStr,
	[Parameter(
		Position = 1)][Alias('Tw')]
		[UInt32]$TabWidth = 8)

$Block = ''
$EOL = $False
$Lines = $InputStr -split "`r`n"
	
foreach($Line in $Lines){
	while($EOL -eq $False){
		$I = $Line.IndexOf("`t")
		if($I -eq -1){$EOL = $True;break}
		$Pad = if($TabWidth -gt 0){
			' ' * ($TabWidth - ($I % $TabWidth))}
		else{''}
		$Line = $Line -replace "^([^\t]{$I})\t(.*)$","`$1$Pad`$2"}
	$Block += '{0}{1}' -f $Line,"`r`n"
	$EOL = $False}
return $Block
}

function Remove-TempFile{
	param(
		[Parameter(Position = 0,Mandatory,ValueFromPipeline)][String]$Path,
		[Parameter(Position = 1)][Alias('DS')][Int]$DelaySeconds = 5)
	#Do-While to force wait for large files
	Do{ Start-Sleep -Seconds $DelaySeconds
		Remove-Item -LiteralPath $Path -ErrorAction SilentlyContinue}
	While(Test-Path -LiteralPath $Path)
}

function Validate-ScriptType{
	param([Int]$Index)
	switch($Index){
		0	{if($ScriptPath -NotMatch $MatchStr){$E=1}}
		1	{$E=1}
	}
	if($E -eq 1){
		Write-Error -Exception (-join ($My.Name,$ErrMsg[$Index]))
		Exit(1)}
}

function Select-InputFile{
	if($Path.Length -eq 0){
		switch($Host.Name){
			$PSIseHost{
				$Script:ScriptName = $psISE.CurrentFile.DisplayName
				$Script:ScriptPath = $psISE.CurrentFile.FullPath
				Validate-ScriptType 0
				$Script:Text = if($psISE.CurrentFile.Editor.SelectedText.Length -eq 0){
							$psISE.CurrentFile.Editor.Text}
						else{
							$psISE.CurrentFile.Editor.SelectedText}}
			$VSCodeHost{
				$PSEContext = $PSEditor.GetEditorContext()
				$Script:ScriptPath = $PSEContext.CurrentFile.Path
				$Script:ScriptName = Split-Path -Path $ScriptPath -Leaf
				Validate-ScriptType 0
				$Script:Text = $PSEContext.CurrentFile.GetText()}}
	}
	elseif($Path){
		$Script:ScriptName = Split-Path -Path $Path -Leaf
		$Script:ScriptPath = $Path
		Validate-ScriptType 0
		$Script:Text = (Get-Content -Path $Path) -join "`r`n"
		}
	else{Validate-ScriptType 1}
}
#endregion

#region Listing Functions
function Append-HtmlSpan($Block, $TokenColor){ 
	if (($TokenColor -eq 'NewLine') -or ($TokenColor -eq 'LineContinuation')){ 
		if($TokenColor -eq 'LineContinuation')
			{[Void]$CodeBuilder.Append('`')}
		[Void]$CodeBuilder.Append("<BR/>`r`n") 
		[Void]$LineBuilder.Append($LineNoFormat -f $LineNo) 
		$Script:LineNo++}
	else{
		$Block = [System.Web.HttpUtility]::HtmlEncode($Block) 
		if(-not $Block.Trim()){$Block = $Block.Replace(' ','&nbsp;')}
		$HtmlColor = [Drawing.Color]::FromKnownColor($TokenColors[$TokenColor]).Name
		if($TokenColor -eq 'String' -or $TokenColor -eq 'Comment' ){
			$Lines = $Block -split "`r`n"
			$Block = ""
			$MultipleLines = $False 
			foreach($Line in $Lines){ 
				if($MultipleLines){
					$Block += "<BR/>`r`n"
					[Void]$LineBuilder.Append($LineNoFormat -f $LineNo)
					$Script:LineNo++}
				$NewText = $Line.TrimStart()
				$NewText = "&nbsp;" * ($Line.Length - $NewText.Length) + $NewText
				if($TokenColor -eq 'Comment'){$NewText = $NewText.Replace(' ','&nbsp;')}
				$Block += $NewText
				$MultipleLines = $True}}
		[Void]$CodeBuilder.Append("<span style='color:$HtmlColor'>$Block</span>")}
}
 
function Get-HtmlClipboardFormat($Html,$Caller){
	$Header = @"
Version:${AppVersion}
StartHTML:0000000000
EndHTML:0000000000
StartFragment:0000000000
EndFragment:0000000000
StartSelection:0000000000
EndSelection:0000000000
SourceURL:file:///about:blank
<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">
<HTML>
<HEAD> 
<TITLE>HTML Clipboard</TITLE> 
</HEAD> 
<BODY> 
<!--StartFragment--> 
<DIV CLASS='footerstyle'></DIV>
<DIV style='$HtmlFont 
	width:950; border:0px solid black; padding:5px'> 

<TABLE BORDER='0' cellpadding='5' cellspacing='0'> 
<TR> 
	<TD VALIGN='Top'> 
<DIV style='$HtmlFont 
	padding:5px; background:#cecece'> 
__LINES__ 
</DIV> 
	</TD> 
	<TD VALIGN='Top' NOWRAP='NOWRAP'> 
<DIV style='$HtmlFont 
	padding:5px; background:#fcfcfc'> 
__HTML__ 
</DIV> 
	</TD> 
</TR> 
</TBODY> 
</TABLE> 
</DIV> 
<!--EndFragment--> 
</BODY> 
</HTML> 
"@
	$SF = '<!--StartFragment-->'
	$HtmlTag = '__HTML__'
	$Header = $Header.Replace('__LINES__', $LineBuilder.ToString())
	$StartFragment = $Header.IndexOf($SF) + $SF.Length + 2
	$EndFragment = $Header.IndexOf('<!--EndFragment-->') +
		$Html.Length - $HtmlTag.Length
	$StartHtml = $Header.IndexOf('<!DOCTYPE')
	$EndHtml = $Header.Length + $Html.Length - $HtmlTag.Length

	$HRV = @('StartHTML:';'EndHTML:';'StartFragment:';'EndFragment:';'StartSelection:';'EndSelection:')
	$DS = '0'*10
	$FS = '{0:D10}'

	if ($Caller -eq 'Print'){
		$Header = $Header.Replace("Version:${AppVersion}",'')
		0..5|ForEach-Object{$Header = $Header.Replace($HRV[$_]+$DS,'')}
		$Header = $Header.Replace('SourceURL:file:///about:blank','')}
	else{
		$Header = $Header -replace ($HRV[0]+$DS), ($HRV[0]+$FS -f $StartHtml) 
		$Header = $Header -replace ($HRV[1]+$DS), ($HRV[1]+$FS -f $EndHtml) 
		$Header = $Header -replace ($HRV[2]+$DS), ($HRV[2]+$FS -f $StartFragment) 
		$Header = $Header -replace ($HRV[3]+$DS), ($HRV[3]+$FS -f $EndFragment) 
		$Header = $Header -replace ($HRV[4]+$DS), ($HRV[4]+$FS -f $StartFragment) 
		$Header = $Header -replace ($HRV[5]+$DS), ($HRV[5]+$FS -f $EndFragment)}

	$SV = 'HTML Clipboard'
	if($Path.Length -eq 0){
		switch($Host.Name){
			$psISEHost{$Header = $Header.Replace($SV,$psISE.CurrentFile.DisplayName)}
			$VSCodeHost{$Header = $Header.Replace($SV,$ScriptName)}}}
	else{$Header = $Header.Replace($SV,(Split-Path -Path $Path -Leaf))}
	$Header = $Header.Replace($HtmlTag, $Html)
	Write-Verbose -Message $Header
	$Header
}

function Create-PSListing{
	$Script:LineNo = 1
	Select-InputFile
	#Tab Expansion
	$Text = $Text|Expand-Tabs -Tw $TabWidth
	#Do Syntax Parsing.
	$Tokens = [System.Management.Automation.PsParser]::Tokenize($Text, [ref] $Null)
	#Initialize HTML Builder.
	$CodeBuilder = New-Object -TypeName System.Text.StringBuilder
	$LineBuilder = New-Object -TypeName System.Text.StringBuilder
	[Void]$LineBuilder.Append($LineNoFormat -f $LineNo)
	$Script:LineNo++
	#Iterate over the tokens and set their colors.
	$Position = 0
	foreach($Token in $Tokens){
		if ($Position -lt $Token.Start){
			#Second
			$Block = $Text.Substring($Position, ($Token.Start - $Position))
			$TokenColor = 'Unknown'
			Append-HtmlSpan $Block $TokenColor}
		#First
		$Block = $Text.Substring($Token.Start, $Token.Length)
		$TokenColor = $Token.Type.ToString()
		Append-HtmlSpan $Block $TokenColor
		$Position = $Token.Start + $Token.Length
	}

	$R = @('.psm*1','.htm','\*$')
	$Html = Get-HtmlClipboardFormat $CodeBuilder.ToString() 'Print'
	if($OpenInWebBrowser.IsPresent){
		if($OutputToTemp.IsPresent){
		$Script:OutName = Join-Path -Path $Env:Temp -ChildPath ($ScriptName -replace($R[0],$R[1]) -replace $R[2],'')}
		else{
			$Script:OutName = ($ScriptPath -replace($R[0],$R[1]) -replace $R[2],'')}
			$Html | Out-File -FilePath $OutName
			if($My.ModeIdx -eq [ModeOfOperation]::Listing){
				Invoke-Item -LiteralPath $OutName
				if($OutputToTemp.IsPresent){$OutName|Remove-TempFile}
			}
	}
	else{
		$OutName = $ScriptPath -replace($R[0],$R[1])
		$Html | Out-File -FilePath $OutName}
}
#endregion

#region PsToken Report
function Set-OutputFileName{
	$I = $My.OutputIdx + 1
	if($I -eq [TokenReportOutput]::File + 1 -or $I -eq [TokenReportOutput]::Html + 1){
		$Script:OutName = if($OutputToTemp.IsPresent){
	Join-Path -Path $Env:Temp -ChildPath ($ScriptName -replace($RVStr[0],$RVStr[$I]) -replace $RVStr[2], $RVStr[3])}
	else{
		($ScriptPath -replace($RVStr[0],$RVStr[$I]) -replace $RVStr[2], $RVStr[3])}
	}
}

function Build-TokenList{
	if($My.OutputIdx -ne [TokenReportOutput]::Raw){
		#Filter to Selected PsToken Types
		$TokenTypes = @('Command','CommandArgument','CommandParameter','Member','Type','Unknown','Variable')
		if($ShowComments.IsPresent){$TokenTypes += 'Comment';[Array]::Sort($TokenTypes)}
		$Tokens = $RawTokens|Where-Object{($TokenTypes.Contains($_.Type.ToString()))}

		#Build Custom Output Objects & Sort
		$Script:Tokens = $Tokens|ForEach-Object{
			$NewObj = [PSCustomObject][Ordered]@{
				Type = $_.Type.ToString() 
				Content = $(
					if($_.Type -eq 'Comment'){
						$_.Content.Substring(0,$(
						if($_.Content.Length -gt $MaxContentLength){$MaxContentLength}
						else{$_.Content.Length-1}))}
					else{$_.Content})
				StartLine = $_.StartLine
				StartColumn = $_.StartColumn
				EndLine = $_.EndLine
				EndColumn = $_.EndColumn
				Length = $_.Length}
			if($ShowStart.IsPresent){
				Add-Member -InputObject $NewObj -MemberType NoteProperty -Name 'Start' -Value $_.Start}
			return $NewObj}|Sort-Object -Property Type,Content,Startline
	}
}

function Write-ToFile{
	$Tokens|Format-Table -AutoSize | Out-File -Filepath ($OutName)
	if($OutputToTemp.IsPresent){
		switch($Host.Name){
			$PSIseHost	{$psISE.CurrentPowerShellTab.Files.Add($OutName)}
			$VSCodeHost {$PSEditor.Workspace.OpenFile($OutName)}
				default {NotePad.exe $OutName}}
		$OutName|Remove-TempFile}
}

function Write-ToHtml{
	$Hdr = -join ('&nbsp;' * 13),'<b>',$ScriptName,'-',$ReportName,'</b>'+('<br/>' * 2)
	$Html = $Tokens|ConvertTo-Html -As Table -Fragment -Property *
	ConvertTo-Html -Body $Html -Title $ReportName -Head $Hdr| Out-File $OutName -Append:($My.ModeIdx -eq [ModeOfOperation]::Both)
	if($OpenInWebBrowser.IsPresent){Invoke-Item -LiteralPath $OutName}
		if($OutputToTemp.IsPresent){$OutName|Remove-TempFile}
}

function Write-TokenResults{
	switch($My.OutputIdx){
		([Int][TokenReportOutput]::File)		{Write-ToFile}
		([Int][TokenReportOutput]::GridView)	{$Tokens|Out-GridView -Title "$ScriptName - $ReportName"}
		([Int][TokenReportOutput]::Host)		{$Tokens|Format-Table -AutoSize}
		([Int][TokenReportOutput]::Html)		{Write-ToHtml}
		([Int][TokenReportOutput]::Raw)			{$RawTokens}
	}
}

function Create-TokenReport{
	Select-InputFile
	#Tab Expansion
	$Script:Text = $Text|Expand-Tabs -Tw $TabWidth
	#Tokenize Input Script
	$Script:RawTokens = [System.Management.Automation.PsParser]::Tokenize($Text, [Ref] $Null)
	Build-TokenList
	if($My.ModeIdx -ne [ModeOfOperation]::Both){Set-OutputFileName}
	Write-TokenResults
}
#endregion

#region Report Output by Mode
Switch($My.ModeIdx){
	([Int][ModeOfOperation]::Listing)	{Create-PSListing}
	([Int][ModeOfOperation]::Tokens)	{Create-TokenReport}
	([Int][ModeOfOperation]::Both)		{Create-PSListing; Create-TokenReport}
}
#endregion

#region Test Commands
<#
.\Create-PSListing.ps1 -Mode Listing -LW 3 -Out Html -ICT -SS -T -O
.\Create-PSListing.ps1 -Mode Tokens -Out Html -ICT -SS -T -O
.\Create-PSListing.ps1 -Mode Both -LW 3 -Out Html -ICT -SS -T -O 
#>
#endregion