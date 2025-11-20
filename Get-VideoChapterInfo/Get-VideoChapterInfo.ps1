<#
.NOTES
	File Name:	Get-VideoChapterInfo.ps1
	Version:	Version 1.0 - 2025-11-13
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	2025-11-13

.SYNOPSIS
	Extracts Chapter MetaData from m4v\mp4\mkv files using FFProbe.

.DESCRIPTION
	Extracts Chapter MetaData from m4v\mp4\mkv files using FFProbe.
	The ID reported is the Chapter# not the MetaData ID value as the
	various video file formats use this field as an internal index.
	StartMs, EndMs, & Length are TimeBase Adjusted to Milliseconds.
	FFProbe (A part of FFMpeg) must be installed on the System Path.
	Download FFMpeg for free at: https://ffmpeg.org/download.html

.PARAMETER FullName Alias: F
	Required, Fully qualified video file path to an m4v\mp4 file.

.PARAMETER MillisecondsToTime Alias: T
	Optional, Switch if present causes time values to be
	converted to dd:hh:mm:ss.fff formatted strings.

.INPUTS
	Path of file to examine,
	Output from FFProbe as a process.

.OUTPUTS
	Chapter Information as a CustomPSObject

.EXAMPLE
	PS> .\Get-VideoChapterInfo.ps1 -FullName "\\MyCloud2\Public\Shared Videos\Movies\Downloads\Without_a_Clue_(1988).mp4"|ft
	Output a list of Chapter Objects as a table with values as Times.

	
	ID TimeBaseNum TimeBaseDen	 Start	   End StartMs   EndMs Length Title		File
	-- ----------- -----------	 -----	   --- -------   ----- ------ -----		----
	 1		   1	   30000		 0  16170000	   0  539000 539000 "Chapter 1"  \\MyCloud2\Public\Shared Videos\Movies\Downloads\Without_a_Clue_(1988).mp4
	 2		   1	   30000  16170000  31963200  539000 1065440 526440 "Chapter 2"
	 3		   1	   30000  31963200  47990400 1065440 1599680 534240 "Chapter 3"
	 4		   1	   30000  47990400  53194800 1599680 1773160 173480 "Chapter 4"
	 5		   1	   30000  53194800  65035200 1773160 2167840 394680 "Chapter 5"
	 6		   1	   30000  65035200  87354000 2167840 2911800 743960 "Chapter 6"
	 7		   1	   30000  87354000  90692400 2911800 3023080 111280 "Chapter 7"
	 8		   1	   30000  90692400 101419200 3023080 3380640 357560 "Chapter 8"
	 9		   1	   30000 101419200 114446400 3380640 3814880 434240 "Chapter 9"
	10		   1	   30000 114446400 124135200 3814880 4137840 322960 "Chapter 10"
	11		   1	   30000 124135200 132621600 4137840 4420720 282880 "Chapter 11"
	12		   1	   30000 132621600 147241200 4420720 4908040 487320 "Chapter 12"
	13		   1	   30000 147241200 165697200 4908040 5523240 615200 "Chapter 13"
	14		   1	   30000 165697200 175776000 5523240 5859200 335960 "Chapter 14"
	15		   1	   30000 175776000 179534400 5859200 5984480 125280 "Chapter 15"
	16		   1	   30000 179534400 183988800 5984480 6132960 148480 "Chapter 16"
	17		   1	   30000 183988800 187160970 6132960 6238699 105739 "Chapter 17"
	18		   1	   30000 187160970 191736540 6238699 6391218 152519 "Chapter 18"

.EXAMPLE
	PS> .\Get-VideoChapterInfo.ps1 -F "\\MyCloud2\Public\Shared Videos\Movies\Downloads\Without_a_Clue_(1988).mp4" -T|ft
	Output a list of Chapter Objects as a table.

	ID TimeBaseNum TimeBaseDen Start		  End			StartMs	  EndMs		Length	   Title		File
	-- ----------- ----------- -----		  ---			-------	  -----		------	   -----		----
	 1		   1	   30000 00:00:00.000   04:29:30.000   00:00:00.000 00:08:59.000 00:08:59.000 "Chapter 1"  \\MyCloud2\Public\Shared Videos\Movies\Downloa...
	 2		   1	   30000 04:29:30.000   08:52:43.200   00:08:59.000 00:17:45.440 00:08:46.440 "Chapter 2"
	 3		   1	   30000 08:52:43.200   13:19:50.400   00:17:45.440 00:26:39.680 00:08:54.240 "Chapter 3"
	 4		   1	   30000 13:19:50.400   14:46:34.800   00:26:39.680 00:29:33.160 00:02:53.480 "Chapter 4"
	 5		   1	   30000 14:46:34.800   18:03:55.200   00:29:33.160 00:36:07.840 00:06:34.680 "Chapter 5"
	 6		   1	   30000 18:03:55.200   1:00:15:54.000 00:36:07.840 00:48:31.800 00:12:23.960 "Chapter 6"
	 7		   1	   30000 1:00:15:54.000 1:01:11:32.400 00:48:31.800 00:50:23.080 00:01:51.280 "Chapter 7"
	 8		   1	   30000 1:01:11:32.400 1:04:10:19.200 00:50:23.080 00:56:20.640 00:05:57.560 "Chapter 8"
	 9		   1	   30000 1:04:10:19.200 1:07:47:26.400 00:56:20.640 01:03:34.880 00:07:14.240 "Chapter 9"
	10		   1	   30000 1:07:47:26.400 1:10:28:55.200 01:03:34.880 01:08:57.840 00:05:22.960 "Chapter 10"
	11		   1	   30000 1:10:28:55.200 1:12:50:21.600 01:08:57.840 01:13:40.720 00:04:42.880 "Chapter 11"
	12		   1	   30000 1:12:50:21.600 1:16:54:01.200 01:13:40.720 01:21:48.040 00:08:07.320 "Chapter 12"
	13		   1	   30000 1:16:54:01.200 1:22:01:37.200 01:21:48.040 01:32:03.240 00:10:15.200 "Chapter 13"
	14		   1	   30000 1:22:01:37.200 2:00:49:36.000 01:32:03.240 01:37:39.200 00:05:35.960 "Chapter 14"
	15		   1	   30000 2:00:49:36.000 2:01:52:14.400 01:37:39.200 01:39:44.480 00:02:05.280 "Chapter 15"
	16		   1	   30000 2:01:52:14.400 2:03:06:28.800 01:39:44.480 01:42:12.960 00:02:28.480 "Chapter 16"
	17		   1	   30000 2:03:06:28.800 2:03:59:20.970 01:42:12.960 01:43:58.699 00:01:45.739 "Chapter 17"
	18		   1	   30000 2:03:59:20.970 2:05:15:36.540 01:43:58.699 01:46:31.218 00:02:32.519 "Chapter 18"
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory)][Alias('F')][String]$FullName,
	[Parameter()][Alias('T')][Switch]$MillisecondsToTime
)

# Define the Chapter class
Class Chapter {
	[int]  $ID
	[long] $TimeBaseNum
	[long] $TimeBaseDen
	[long] $Start
	[long] $End
	[long] $StartMs
	[long] $EndMs
	[long] $Length
	[string] $Title
	[string] $File

	Chapter() {
		$this.ID = 0
		$this.TimeBaseNum = 0
		$this.TimeBaseDen = 0
		$this.Start = 0
		$this.End = 0
		$this.StartMs = 0
		$this.EndMs = 0
		$this.Length = 0
		$this.Title = ''
		$this.File = ''
	}

	Chapter([int]$ID, [long]$Num, [long]$Den, [long]$Start, [long]$End, [string]$Title, [string]$File) {
		$this.ID = $ID
		$this.TimeBaseNum = $Num
		$this.TimeBaseDen = $Den
		$this.Start = $Start
		$this.End = $End
		$this.StartMs = $this.ApplyTimeBase($this.Start)
		$this.EndMs = $this.ApplyTimeBase($this.End)
		$this.Length = $this.EndMs - $this.StartMs
		$this.Title = $Title
		$this.File = $File
	}

	[long] ApplyTimeBase([long]$Value) {
		if ($this.TimeBaseNum -eq 0 -or $this.TimeBaseDen -eq 0) {
			throw "TimeBase numerator/denominator cannot be zero."
		}
		$seconds = $Value * ($this.TimeBaseNum / [double]$this.TimeBaseDen)
		return [long]($seconds * 1000) # milliseconds
	}
}

# Define the ChapterTime class
Class ChapterTime {
	[int]$ID
	[long]$TimeBaseNum
	[long]$TimeBaseDen
	[String]$Start
	[String]$End
	[String]$StartMs
	[String]$EndMs
	[String]$Length
	[String]$Title
	[String]$File

	ChapterTime() {
		$this.ID = 0
		$this.TimeBaseNum = 0
		$this.TimeBaseDen = 0
		$this.Start = ''
		$this.End = ''
		$this.StartMs = ''
		$this.EndMs = ''
		$this.Length = ''
		$this.Title = ''
		$this.File = ''
	}

	ChapterTime([int]$ID, [long]$TimeBaseNum, [long]$TimeBaseDen, [long]$Start, [long]$End, [long]$StartMs, [long]$EndMs, [long]$Length, [string]$Title, [string]$File) {
		$this.ID = $ID
		$this.TimeBaseNum = $TimeBaseNum
		$this.TimeBaseDen = $TimeBaseDen
		$this.Start = Convert-MillisecondsToTimeString($Start)
		$this.End = Convert-MillisecondsToTimeString($End)
		$this.StartMs = Convert-MillisecondsToTimeString($StartMs)
		$this.EndMs = Convert-MillisecondsToTimeString($EndMs)
		$this.Length = Convert-MillisecondsToTimeString($Length)
		$this.Title = $Title
		$this.File = $File
	}
}

# Logging helper
Function Write-Log {
	param([string]$Message)
	Write-Verbose "[Get-VideoChapterInfo] $Message"
}
Function Convert-MillisecondsToTimeString {
	[CmdletBinding()]
	param ([Parameter(Mandatory)][ValidateRange(0, [double]::MaxValue)][double]$Milliseconds)

	try {
		$maxMs = [TimeSpan]::MaxValue.TotalMilliseconds - 1
		if ($Milliseconds -gt $maxMs) {
			throw [System.ArgumentOutOfRangeException]::new('Milliseconds',
				"Value must be between 0 and $([long]$maxMs).")
		}

		$ts = [TimeSpan]::FromMilliseconds($Milliseconds)

		if ($ts.Days -gt 0) {
			return '{0}:{1:00}:{2:00}:{3:00}.{4:000}' -f $ts.Days, $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds
		}
		else {
			return '{0:00}:{1:00}:{2:00}.{3:000}' -f $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds
		}
	}
	catch {
		Write-Error "Error converting milliseconds: $($_.Exception.Message)"
	}
}
Function Invoke-FFProbe {
	param ([Parameter(Mandatory)][string] $VideoFilePath)

	if (-not (Test-Path -LiteralPath $VideoFilePath)) {
		throw "File not found: $VideoFilePath"
	}

	try {
		$process = [System.Diagnostics.Process]::New()
		$process.StartInfo = [System.Diagnostics.ProcessStartInfo]::New()
		$process.StartInfo.FileName = "ffprobe"
		$process.StartInfo.Arguments = "-i `"$VideoFilePath`" -print_format flat -show_chapters -loglevel error"
		$process.StartInfo.RedirectStandardOutput = $true
		$process.StartInfo.RedirectStandardError = $true
		$process.StartInfo.UseShellExecute = $false
		$process.StartInfo.CreateNoWindow = $true

		$null = $process.Start()
		$output = $process.StandardOutput.ReadToEnd()
		$errorOutput = $process.StandardError.ReadToEnd()
		$process.WaitForExit()

		if ($process.ExitCode -ne 0) {
			throw "ffprobe failed with exit code $($process.ExitCode): $errorOutput"
		}

		Write-Log "FFProbe executed successfully."
		return $output
	}
	catch {
		throw "Error running ffprobe: $($_.Exception.Message)"
	}
}
Function Parse-ChapterLines {
	param (
		[Parameter(Mandatory)][string[]]$Lines,
		[Parameter(Mandatory)][int]$Name)

	$props = @{
		ID = $null
		TimeBaseNum = $null
		TimeBaseDen = $null
		Start = $null
		End = $null
		Title = ''
		File = $null
	}

	foreach ($line in $Lines) {
		switch -Regex ($line) {
			'^time_base="(\d+)/(\d+)"$' { $props.TimeBaseNum = [long]$Matches[1]
										  $props.TimeBaseDen = [long]$Matches[2] }
			'^start=(\d+)$'				{ $props.Start	= [long]$Matches[1] }
			'^end=(\d+)$'				{ $props.End	= [long]$Matches[1] }
			'^tags.title=(.+)$'			{ $props.Title	= $Matches[1] }
			default { } # ignore unknown tags
		}
	}

	if ($null -eq $props.TimeBaseNum -or
		$null -eq $props.TimeBaseDen -or
		$null -eq $props.Start -or
		$null -eq $props.End){
		throw "Incomplete chapter metadata: $($Lines -join ', ')"
	}
	$props.ID = $Name + 1 # Chapter No.
	$props.File = if($props.ID -eq 1){$FullName}

	return [Chapter]::new(
		$props.ID,
		$props.TimeBaseNum,
		$props.TimeBaseDen,
		$props.Start,
		$props.End,
		$props.Title,
		$props.File
	)
}
Function Get-Chapters {
	param([string[]]$Data)

	$Chapters = [System.Collections.Generic.List[Chapter]]::new()
	$Groups = $Data | Group-Object { ($_ -split '\.')[0] } -ErrorAction Stop

	foreach ($Group in $Groups) {
		$lines = $Group.Group | ForEach-Object { ($_ -split '\.',2)[1] }
		$chapter = Parse-ChapterLines -Lines $lines -Name $Group.Name
		$Chapters.Add($chapter)
	}

	Write-Log "Parsed $($Chapters.Count) chapters."
	return $Chapters
}
Function Get-MetadataLines {
	$reader = New-Object System.IO.StringReader($FFPData)
	$Lines = [System.Collections.ArrayList]::New()
	try {
		while ($line = $reader.ReadLine()) {
			if ([string]::IsNullOrWhiteSpace($line)) { continue }
			[void]$Lines.Add($line.Trim())
		}
	}
	finally {
		$reader.Dispose()
	}
	return $Lines
}
Function Test-ValidPath(){
Param([Parameter(Mandatory)][Alias('P')][String]$Path)
	# Validate Path
	$VFI = [IO.FileInfo]::new($Path)
	If(!$VFI.Exists){
		Throw [System.IO.FileNotFoundException]::new(
				"The file {0} was not found." -f $VFI.FullName,
				$VFI.FullName)
	}
	if ($VFI.Extension -notin @('.m4v','.mp4','.mkv')) {
	throw [System.ArgumentOutOfRangeException]::new(
		'FullName',
		"Parameter must be an m4v, mp4 or mkv file. {$VFI.FullName}")
}

}
Function Select-Output(){
	if($MillisecondsToTime){
		$ChapterTimes = [System.Collections.Generic.List[ChapterTime]]::new()
															forEach($item in $Chapters){
		$CT = [ChapterTime]::new(
			$item.ID,
			$item.TimeBaseNum,
			$item.TimeBaseDen,
			$item.Start,
			$item.End,
			$item.StartMs,
			$item.EndMs,
			$item.Length,
			$item.Title,
			$item.File)
		$ChapterTimes.Add($CT)
	}
	}
	return $(if(-not $MillisecondsToTime) { $Chapters } else { $ChapterTimes })
}
#Mainline Process
Test-ValidPath -Path $FullName
$Chapters = [System.Collections.Generic.List[Chapter]]::new()
$FFPData = Invoke-FFProbe -VideoFilePath $FullName
$FFPData = $FFPData.replace('chapters.chapter.','')
$Lines = Get-MetadataLines($FFPData)
$Chapters = Get-Chapters -Data $Lines
Select-Output
#End-of-Mainline

#region [Sample FFProbe Flat Chapter items]
<#
Uses Zero-Based ChapterIndex
Path>ChapterIndex>Field=Value
-------------------------------------
chapters.chapter.0.id=0
chapters.chapter.0.time_base="1/1000"
chapters.chapter.0.start=0
chapters.chapter.0.start_time="0.000000"
chapters.chapter.0.end=272798
chapters.chapter.0.end_time="272.798000"
chapters.chapter.0.tags.title="Main Titles\\Opening"
#>
#endregion
