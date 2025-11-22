<#
.NOTES
-------------------------------------
Name:    Convert-SpacesToTabs.ps1
Version: 3.0 - 11/21/2025
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
--------------------------------------------------------------------------------------------
Revision History:
V3.0 - 11/21/2025 Refactored for modular helpers, expanded encoding options, improved stats
				  Replaced 'Default' with explicit UTF8NoBOM, clarified help text
V2.0 - 03/27/2021 Added ability to accept piped input
V1.0 - 07/04/2020 Initial Release
--------------------------------------------------------------------------------------------

.SYNOPSIS
Converts leading spaces to tabs. If FileOut is omitted, creates a working output of FileIn 
with a ".tvx" extension and renames it to FileIn on completion.

.DESCRIPTION
Converts leading spaces to tabs, saving disk space & reducing the number of clusters required.
By default it also converts Unicode UTF-16 encoded files to UTF-8 No BOM. Windows encodes files
as UTF-16 for multi-language support, but UTF-8 is more efficient for English language files.
An option is provided for encoding in UTF-16.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script. PowerShell security may require you run the
Unblock-File cmdlet with the fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 

.Parameter FullName - Alias: In
Required. Input file path. May come from pipeline.

.Parameter FileOut - Alias: Out
Optional. Output file path. If omitted defaults to FileIn with the extension ".tvx". 
Ignored if input is via the pipeline and the AutoNaming option is forced.

.Parameter Encoding - Alias: ES
Sets output file encoding scheme. Valid values:
- UTF8NoBOM (default) > UTF-8 without BOM
- UTF8 > UTF-8 with BOM
- Unicode > UTF-16
- ASCII > ASCII

.Parameter SpacesPerTab - Alias: SPT
Optional. Number of spaces to be replaced by a single tab. Default: 4

.Parameter AllSpaces - Alias: A
Optional. Switch causes conversion of all space groups to be replaced.

.Parameter ShowStats - Alias: R
Optional. Switch causes conversion stats to be output.

.EXAMPLE
PS> .\Convert-SpacesToTabs.ps1 -FullName .\AudioPlayer.ps1 -ShowStats | FT
Creates an .\AudioPlayer.ps1 file in UTF-8 No BOM where leading spaces have
been replaced by tabs in a 4:1 ratio.
This can significantly reduce the file size as seen below:

File         SizeIn SizeOut SizeDiff PercentSaved
----         ------ ------- -------- ------------
AudioPlayer   99786   45923   -53863        53.98

.EXAMPLE
PS> Get-ChildItems -Path F:\SamplePath\*.vb | .\Convert-SpacesToTabs.ps1 -ShowStats
Creates a new file in UTF-8 No BOM for each vb source file where leading spaces
have been replaced by tabs in a 4:1 ratio.
The conversion from Unicode to UTF-8 cuts file size significantly, 16-bit\8-bit chars.
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[Alias('In')][String[]]$FullName,
	[Parameter()][Alias('Out')][String]$FileOut = "",
	[Parameter()][Alias('SPT')][Int32]$SpacesPerTab = 4,
	[Parameter()][Alias('ES')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('UTF8NoBOM','Unicode','UTF8','ASCII')]
		[String]$Encoding = 'UTF8NoBOM',
	[Parameter()][Alias('A')][Switch]$AllSpaces,
	[Parameter()][Alias('R')][Switch]$ShowStats
)

Begin {
	$Fi = @($null,$null)

	# Regex caches for compiled patterns
	$LeadingRegexCache = @{}
	$AllRegexCache     = @{}

	function Convert-LeadingSpaces {
		param([string]$Line,[int]$SpacesPerTab)
		if (-not $LeadingRegexCache.ContainsKey($SpacesPerTab)) {
			$pattern = "^(?: {$SpacesPerTab})+"
			$LeadingRegexCache[$SpacesPerTab] = [Regex]::new($pattern,[System.Text.RegularExpressions.RegexOptions]::Compiled)
		}
		$regex = $LeadingRegexCache[$SpacesPerTab]
		$regex.Replace($Line, { "`t" * ($args[0].Value.Length / $SpacesPerTab) })
	}

	function Convert-AllSpaces {
		param([string]$Line,[int]$SpacesPerTab)
		if (-not $AllRegexCache.ContainsKey($SpacesPerTab)) {
			$pattern = "( {$SpacesPerTab})"
			$AllRegexCache[$SpacesPerTab] = [Regex]::new($pattern,[System.Text.RegularExpressions.RegexOptions]::Compiled)
		}
		$regex = $AllRegexCache[$SpacesPerTab]
		$regex.Replace($Line,"`t")
	}

	function Write-ConversionStats {
		param($FileIn,$FileOut)
		[PSCustomObject]@{
			File         = [IO.Path]::GetFileNameWithoutExtension($FileOut.Name)
			SizeIn       = $FileIn.Length
			SizeOut      = $FileOut.Length
			SizeDiff     = $FileOut.Length - $FileIn.Length
			PercentSaved = [math]::Round((1 - ($FileOut.Length / $FileIn.Length)) * 100,2)
		}
	}

	Enum ConversionMode { AllSpaces; LeadingSpaces }

	function Get-EncodingObject {
		param([string]$EncodingName)
		switch ($EncodingName) {
			'UTF8NoBOM' { [Text.UTF8Encoding]::new($false) }
			'UTF8'      { [Text.UTF8Encoding]::new($true) }
			'Unicode'   { [Text.Encoding]::Unicode }
			'ASCII'     { [Text.Encoding]::ASCII }
			default     { [Text.UTF8Encoding]::new($false) }
		}
	}
}
Process {
	foreach ($FileIn in $FullName) {
		if ($PSCmdlet.MyInvocation.ExpectingInput) {
			$Fi = @($null,$null)
			$FileOut = "" # Force Auto-Naming
		}

		$AutoOut = ($FileOut.Length -eq 0)

		try {
			$Fi[0] = Get-ChildItem -Path $FileIn -ErrorAction Stop
		}
		catch {
			Write-Error -Message "Error: File - $FileIn Not Found! $_" -Category ResourceUnavailable
			continue
		}

		if ($Fi[0].Exists) {
			if ($AutoOut) {
				$FileOut = Join-Path $Fi[0].DirectoryName "$([IO.Path]::GetFileNameWithoutExtension($Fi[0].Name)).tvx"
			}

			$Mode        = if ($AllSpaces) { [ConversionMode]::AllSpaces } else { [ConversionMode]::LeadingSpaces }
			$OutEncoding = Get-EncodingObject -EncodingName $Encoding

			try {
				$reader = [IO.StreamReader]::new($FileIn)
				$writer = [IO.StreamWriter]::new($FileOut, $false, $OutEncoding)

				while (-not $reader.EndOfStream) {
					$line = $reader.ReadLine()
					if ($Mode -eq [ConversionMode]::AllSpaces) {
						$writer.WriteLine((Convert-AllSpaces $line $SpacesPerTab))
					} else {
						$writer.WriteLine((Convert-LeadingSpaces $line $SpacesPerTab))
					}
				}

				$reader.Close()
				$writer.Close()
			}
			catch {
				Write-Error -Message "Error processing file $FileIn : $_"
				continue
			}

			if ($ShowStats) {
				$Fi[1] = Get-ChildItem -Path $FileOut
				Write-ConversionStats -FileIn $Fi[0] -FileOut $Fi[1]
			}

			if ($AutoOut) {
				try {
					$BakFile = "$FileIn.bak"
					if (Test-Path $BakFile) { Remove-Item -LiteralPath $BakFile -ErrorAction SilentlyContinue }
					Rename-Item -Path $FileIn -NewName $BakFile -ErrorAction Stop
					Rename-Item -Path $FileOut -NewName $FileIn -ErrorAction Stop
				}
				catch {
					Write-Error -Message "Error renaming backup/output files for $FileIn : $_"
				}
			}
		}
	}
}
End {}
