<#
.NOTES
	File Name:	Set-VideoMediaLangEx.ps1
	Version:	Version: 1.5 - 2026/04/21
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	2025/06/20
	Revision History:
		V1.5 - 2026/04/21 - Added audio track name support
		V1.4 - 2026/04/19 - Added input deletion switch
		V1.3 - 2025/08/09 - Added ability to set target audio track#
		V1.2 - 2025/08/08 - Added Video Language Dictionary json
		V1.1 - 2025/08/07 - Added redirection support
	 	V1.0 - 2025/06/20 - Original Wersion

.SYNOPSIS
	This script uses FFmpeg to easily set an audio track language of a supported video file(s)
	without rerendering.

.DESCRIPTION
	This script uses FFmpeg to easily set an audio track language of a supported video file(s)
	without rerendering. Input filenames may be input via the pipeline. The resulting output files
	are named the same as the input, but are directed to another directory. The default output directory
	will be created, if necessary. Any existing output files in the output directory are
	overwritten. The FFmpeg audio track language is a 3-character ISO 639-2 Code,
	the default is 'eng' (English). FFmpeg logging is restricted to errors only by default
	and the filename of each processed file is shown instead. You may display full FFmpeg
	logging by use of the -FullLog switch.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------

.COMPONENT
	An installed copy of FFmpeg.exe on the system path.
	An ISO639-2_Video_Language_Codes.json in the same location as the script.

.PARAMETER FullName Alias: FN
	Requried, The fully qualified filename of the input file.

.PARAMETER PathOut Alias: PO
	Optional, The fully qualified output directory path.
	defaults to a 'New' directory below the input directory.

.PARAMETER LanguageID Alias: AL (Default is eng)
	Optional, 3-character ISO 639-2 Language Code.
	See: https://www.loc.gov/standards/iso639-2/php/code_list.php
	A ISO639-2_Video_Language_Codes.json was added in v1.2

.PARAMETER AudioTrack Alias: AT (Default is 0)
	Optional, Zero-based audio track# to modify.

.PARAMETER AudioTrackName Alias: AN
	Optional Audio Track Title Default is 'Primary'

.PARAMETER NameAudio Alias: NA
	Optional, This switch enables Audio Track Naming
	
.PARAMETER FullLog Alias: FL
	Optional, This switch will cause full FFmpeg logging to be shown.

.PARAMETER DelIn Alias: DI
	Optional, This switch will cause input file deletion.
	
.INPUTS
	The fully qualified filename of the input file.

.OUTPUTS
	Updated video files.	

.EXAMPLE
	PS> .\Set-VideoMediaLang.ps1 -FullName f:\SuperCar-1961\SuperCar-1961_S1E01_Italy.mkv =LanguageID ita -FullLog
	This command uses one input file and directs output to a "New" subfolder. The first audio track is set to
	Italian. The full text from FFmpeg is displayed in the console

.EXAMPLE
	PS> (GCI -Path f:\SuperCar-1961\*.mkv)|.\Set-VideoMediaLang.ps1 
	This command uses the pipeline to pass a list of input files and directs output to a "New" subfolder.
	Only limited status information is displayed, unless an FFmpeg error is detected.
	Optional,
	PS> (GCI -Path "f:\Video\InternetArchive\SuperCar-1961\*.mkv")|.\Set-VideoMediaLang.ps1 -FullLog|Out-File Set-VideoMediaLang.log
	This command uses the pipeline to pass a list of input files and directs that the full FFmpeg log be saved to a text file.

.EXAMPLE
	PS> .\Set-VideoMediaLang.ps1 -FN f:\Video\Around_the_World_in_80_Days_HD.mp4 -AL fre -AT 1
	This example would set the second audio track language to French.
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage='Fully Qualified Input File Name')][Alias('FN')][String]$FullName,
	[Parameter(HelpMessage='The output target directory')][Alias('PO')][String]$PathOut='',
	[Parameter(HelpMessage='The output target language code')][Alias('AL')][String]$LanguageID = 'eng',
	[Parameter(HelpMessage='The output target audio track(Zero Based)')][Alias('AT')][Int]$AudioTrack = 0,
	[Parameter(HelpMessage='The output target audio track title')][Alias('AN')][String]$AudioTrackName = 'Primary',
	[Parameter()][Alias('NA')][Switch]$NameAudio,
	[Parameter()][Alias('FL')][Switch]$FullLog,
	[Parameter()][Alias('DI')][Switch]$DelIn	
	)

Begin{
#region Variables
$Locked = $False
$LanguageID = $LanguageID.ToLower()
$FFmpeg = 'FFmpeg.exe'
$JsonFile = @{File='.\ISO639-2_Video_Language_Codes.json';Info=$null}
#region Utility Functions
Function Import-VideoLanguageIdDictionary{
Param([Parameter(Mandatory)][Alias('VDF')][String]$VideoDictionaryFile)
	$CodeDict = [System.Collections.Generic.Dictionary[String,String]]::New()
	(Get-Content -Path $VideoDictionaryFile | ConvertFrom-Json) | ForEach-Object{$CodeDict.Add($_.Code,$_.Name)}
	$CodeDict
}
Function Resolve-CurrentLocation{
	param([Parameter(Mandatory)][Alias('P')][String]$Path)
	return "{0}\{1}" -f (Get-Location),(Split-Path -Path $($Path) -Leaf)
}
#endregion
#region Startup Code
	if(!(Get-Command $FFmpeg -ErrorAction SilentlyContinue)){throw ('{0} not found on PATH.' -f $FFmpeg)}
	$JsonFile.Info = [IO.FileInfo]::new((Resolve-CurrentLocation -Path $JsonFile.File))
	if(!$JsonFile.Info.Exists){Throw [System.IO.FileNotFoundException] ('JSON Language File: ({0}) Not Found!' -f $JsonFile.Info.FullName)}
	$VideoLanguageIdDictionary = Import-VideoLanguageIdDictionary -VDF $JsonFile.Info.FullName
	if(!$VideoLanguageIdDictionary.ContainsKey($LanguageID)){
		Throw ('Invalid LanguageID: {0}.' -f $LanguageID)
	}
	if(!$FullLog){
		'Processing ...'
		'Target Language Code: {0} Language: {1} Track#: {2}' -f $LanguageID,$VideoLanguageIdDictionary[$LanguageID],$AudioTrack
	}
#endregion
}
Process{
#region Utility Functions
Function Run-FFmpeg(){
	$Prefix = '-hide_banner -loglevel error '
	$MetaCmd = '-metadata:s:a:{0}' -f $AudioTrack
	$CmdLn = '-map 0 -c:a copy -c:v copy {0}' -f $MetaCmd
	$Masks = [Ordered]@{
		LangOnly = '{0} "{1}" {2} {3}{4} "{5}"'
		LangName = '{0} "{1}" {2} {3}{4} {5} {6}"{7}" "{8}"'}
	# NOTE: MaskArgs.* must remain [Ordered] to preserve format-string parameter order.		
	$MaskArgs = [Ordered]@{
		LangOnly = [Ordered]@{
		    InputSwitch   = '-y -i'
		    InputFile     = $FI.FullName
		    CmdLine       = $CmdLn
		    LangKey       = 'language='
		    LangValue     = $LanguageID
		    OutputFile    = Join-Path $PathOut $FI.Name
		}
		LangName = [Ordered]@{
		    InputSwitch   = '-y -i'
		    InputFile     = $FI.FullName
		    CmdLine       = $CmdLn
		    LangKey       = 'language='
		    LangValue     = $LanguageID
		    TitleMetaCmd  = $MetaCmd
		    TitleKey      = 'title='
		    TitleValue    = $AudioTrackName
		    OutputFile    = Join-Path $PathOut $FI.Name
		}
	}
	if(!$NameAudio){
		$StdArgs = $Masks.LangOnly -f $($MaskArgs.LangOnly.Values)}
	else{
		$StdArgs = $Masks.LangName -f $($MaskArgs.LangName.Values)}
	$ArgList = if(!$FullLog){$Prefix+$StdArgs}else{$StdArgs}
	$Process = [System.Diagnostics.Process]::New()
	$Process.StartInfo.FileName = $FFmpeg
	$Process.StartInfo.Arguments = $ArgList
	$Process.StartInfo.RedirectStandardOutput = $False
	$Process.StartInfo.RedirectStandardError = $True
	$Process.StartInfo.UseShellExecute = $False
	$Process.StartInfo.CreateNoWindow = $True
	$Process.Start();
	#All FFmpeg status info is output to StandardError
	[String]$OutputText = $Process.StandardError.ReadToEnd()
	$Process.WaitForExit()
	If($Process.ExitCode -eq 0){
		#Delete Input File
		if($DelIn){$FI.Delete()}
	}
	else{
		#Echo FFmpeg error & Abort
		Throw [PSCustomObject][Ordered]@{ExitCode=$Process.ExitCode;OutputText=$OutputText}} 
	#Return a custom object with Process & OutputText
	return [PSCustomObject][Ordered]@{Process=$Process;OutputText=$OutputText}
}
#endregion
#region	Mainline
	$FI = [IO.FileInfo]::New($FullName)
	if ($PathOut.Length -eq 0){
		$PathOut = [IO.Path]::Combine($FI.Directory,'New')
		if(![IO.Directory]::Exists($PathOut)){
			New-Item -Path $PathOut -ItemType Directory -ErrorAction Stop|Out-Null
		}
	}else{
		$DI = [IO.DirectoryInfo]::New($PathOut)
		if(!$DI.Exists){Throw "PathOut Invalid, $PathOut"}
	}
	if(!$FullLog -and !$Locked){
		'Target Directory: {0}' -f $PathOut
		$Locked = $True
	}
	$RV = Run-FFmpeg
	if($FullLog){('{2}{0} End of File {0}{1}'-f ('-'*50),"`r`n",$RV.OutputText)}Else{$FI.Name}
#endregion
}
End{'--- Process Complete! ---'}

<# Sample\Test commands
(GCI -Path "f:\Video\InternetArchive\SuperCar-1961\*.mkv")|.\Set-VideoMediaLang.ps1 -FullLog
(GCI -Path "f:\Video\InternetArchive\SuperCar-1961\*.mkv")|.\Set-VideoMediaLang.ps1 -LanguageID deu|Out-GridView
(GCI -Path "f:\Video\InternetArchive\SuperCar-1961\*.mkv")|.\Set-VideoMediaLang.ps1 -LanguageID Ita -FullLog|Out-File Set-VideoMediaLang.log
gci c:\Users\RANDY\Documents\DVDFab\DVDFab12\FullDisc\Amazon\FB\*.mp4|.\Set-VideoMediaLang.ps1 -DelIn
.\Set-VideoMediaLangEx3.ps1 -FullName 'c:\Users\RANDY\Documents\DVDFab\DVDFab12\FullDisc\Amazon\FB\Disable Copilot.mp4' -NA -AN 'Main'
#>