<#
.NOTES
	File Name:	Set-VideoMediaLang.ps1
	Version:	Version: 1.1 - 2025/08/07
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	2025/06/20
	Revision History:
		V1.2 - 2025/08/09 - Added Video Language Dictionary json
		V1.1 - 2025/08/07 - Added redirection support
	 	V1.0 - 2025/06/20 - Original Wersion

.SYNOPSIS
	This script uses FFmpeg to easily set the first audio track language of a supported video file(s)
	without rerendering.

.DESCRIPTION
	This script uses FFmpeg to easily set the first audio track language of a supported video file(s)
	without rerendering. Input filenames may be input via the pipeline. The resulting output files
	are named the same as the input, but are directed to another directory. The output directory
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

.PARAMETER FullLog Alias: FL
	Optional, Thi switch will cause full FFmpeg logging to be shown.

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
.EXAMPLE
	PS> (GCI -Path "f:\Video\InternetArchive\SuperCar-1961\*.mkv")|.\Set-VideoMediaLang.ps1 -FullLog|Out-File Set-VideoMediaLang.log
	This command uses the pipeline to pass a list of input files and directs that the full FFmpeg log be saved to a text file.
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage='Fully Qualified Input File Name')][Alias('FN')][String]$FullName,
	[Parameter(HelpMessage='The output target directory')][Alias('PO')][String]$PathOut='',
	[Parameter(HelpMessage='The output target language code')][Alias('AL')][String]$LanguageID = 'eng',
	[Parameter()][Alias('FL')][Switch]$FullLog
	)

Begin{
#region Variables
$Locked = $False
$LanguageID = $LanguageID.ToLower()
$VideoLanguageIdDictionary = $Null
#endregion
#region Utility Functions
Function Import-VideoLanguageIdDictionary{
Param([Parameter(Mandatory)][Alias('VDF')][String]$VideoDictionaryFile)
	$CodeDict = [System.Collections.Generic.Dictionary[String,String]]::New()
	$jsonData = Get-Content -Path $VideoDictionaryFile | ConvertFrom-Json
	$jsonData|ForEach-Object{$CodeDict.Add($_.Code,$_.Name)}
	$CodeDict
}
#endregion
#region Startup Code
	$VideoLanguageIdDictionary = Import-VideoLanguageIdDictionary -VDF .\ISO639-2_Video_Language_Codes.json
	if(!$FullLog){
		'Processing ...'
		'Target Language Code: {0} Language: {1}' -f $LanguageID,$VideoLanguageIdDictionary[$LanguageID]
#endregion
	}
}
Process{
#region Utility Functions
Function Run-FFmpeg(){
	$Prefix = '-hide_banner -loglevel error '
	$StdArgs = '{0} "{1}" {2}{3} "{4}"' -f 
		'-y -i',
		$FI.FullName,
		'-map 0 -c:a copy -c:v copy -metadata:s:a:0 language=',
		$LanguageID,
		$(Join-Path -Path $PathOut -ChildPath $FI.Name)
	$ArgList = if(!$FullLog){$Prefix+$StdArgs}else{$StdArgs}
	$Process = [System.Diagnostics.Process]::New()
	$Process.StartInfo.FileName = 'FFmpeg.exe'
	$Process.StartInfo.Arguments = $ArgList
	$Process.StartInfo.RedirectStandardOutput = $False
	$Process.StartInfo.RedirectStandardError = $True
	$Process.StartInfo.UseShellExecute = $False
	$Process.StartInfo.CreateNoWindow = $True
	$Process.Start();
	#All FFmpeg status info is output to StandardError
	[String]$OutputText = $Process.StandardError.ReadToEnd()
	$Process.WaitForExit()
	If($Process.ExitCode -ne 0){
		#Echo FFmpeg error & Abort
		Throw $OutputText} 
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
#>