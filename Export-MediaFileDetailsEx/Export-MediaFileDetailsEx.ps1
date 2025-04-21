<#
.NOTES
--------------------------------------
Name:    Export-MediaFileDetailsEx.ps1
Version: 1.0i - 05/25/2018
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
--------------------------------------

.SYNOPSIS
This Cmdlet will search a specified location for media files and export 
detailed information about the files to a Tab-Delimited text file suitable 
for importing into a database like MS-Access. This script, an enhanced version 
of my earlier Export-MediaFileDetails.ps1, has been designed using a seperate 
thread to collect the file information, in order to minimize the memory 
requirements. There are 3 modes of operation: Audio, Image, & Video. 
This version has been enhanced to improve performance by 16% over v1.0e.
This version was tested on Windows 10, but should work on others.
    
.DESCRIPTION
Media File Extended File Property Exporter (Multi-Threaded)

.PARAMETER Mode
Mode of Operation <Required> Audio, Image, or Video

.PARAMETER Directory Alias: Dir
Name of the topmost directory to search, if omitted the User default is used.

.PARAMETER OutFile Alias: Out
Name of the export file, if omitted a predefined file of
<Mode>Report.txt is used.

.PARAMETER LogFile Alias: Log
Name of the File Count Log file, if omitted a predefined file of
".\EventTimeLog.txt" is used.

.PARAMETER Append Alias: A
Switch to enable appending to an existing output file.

.PARAMETER Recurse Alias: R
Switch to enable searching recursively.

.PARAMETER RawOutput Alias: RO
Use this switch to overide outputting to the OutFile &
output the Custom PSObjects to the host allowing pipeing to other cmdlets.
The Custom PSObject includes the RunspaceId and PSSourceJobInstanceId of each PSJob thread.

.PARAMETER LogFileCount Alias: L
Switch to enable logging the processed file count to $LogFile.
Overridden by -RawOutput switch.

.PARAMETER LogAppend Alias: LA
Switch to enable appending to the $LogFile.

.PARAMETER Notify Alias: N
Use this switch to have a MessageBox Displayed
upon completion of the selected export operation.

.EXAMPLE
To collect Video File Properties
Export-MediaFileDetailsEx -Mode Video -Dir "\\MediaServer\Public\Video" -Out ".\VideoReport.txt" -R -A

.EXAMPLE
To Display Raw output in a GridView
Export-MediaFileDetailsEx -Mode Video -Dir "\\MediaServer\Public\Video" -R -A -RO|Out-GridView
#>

[CmdletBinding()]
param(   
	[Parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][ValidateSet('Audio','Image','Video')][String]$Mode,
	[Parameter(Mandatory=$False)][Alias('Dir')][String]$Directory="",
	[Parameter(Mandatory=$False)][Alias('Out')][String]$OutFile="",
	[Parameter(Mandatory=$False)][Alias('Log')][String]$LogFile="",
	[Parameter(Mandatory=$False)][Alias('A')][Switch]$Append,
	[Parameter(Mandatory=$False)][Alias('R')][Switch]$Recurse,
	[Parameter(Mandatory=$False)][Alias('RO')][Switch]$RawOutput,
	[Parameter(Mandatory=$False)][Alias('LC')][Switch]$LogFileCount,
	[Parameter(Mandatory=$False)][Alias('LA')][Switch]$LogAppend,
	[Parameter(Mandatory=$False)][Alias('N')][Switch]$Notify)

Import-Module -Name .\MetadataIndexLib.ps1 -Force
Import-Module -Name .\ScriptPathInfo.ps1 -Force

#region Script level variables
$ValidModes = @('Audio','Image','Video')
$ModeIdx = [array]::IndexOf($ValidModes,$Mode)
$Debug = 0
$CurrDate = (Get-Date -DisplayHint Date -Format "yyyy-MM-dd")
$CurrTime = (Get-Date -DisplayHint Time -Format "HH:mm:ss")
$ScriptID = Get-ScriptPathInfo -Mode FileName
if($LogFile.Length -eq 0){$LogFile = -join (".\",$ScriptID,"_Log.txt")}
#endregion

#region Supported File Types
$MediaFileTypes = 
@(
#Audio
	(
	".3gp",".aac",".act",".aiff",".alac",".amr",".atrac",".au",
	".awb",".dct",".dss",".dvf",".flac",".gsm",".iklax",".ivs",
	".m3u",".m4a",".m4p",".mmf",".mp2",".mp3",".mpc",".msv",".ogg",
	".opus",".ra",".rm",".raw",".tta",".vox",".wav",".wavpack",".wma"
	),
#Image
	(
	".bmp",".clp",".emf",".img",".jp2",
	".jpg",".jpeg",".mac",".pcx",".png",
	".gif",".tif",".tiff",".ras",".raw"
	),
#Video
	(
	".avi",".divx",".dvx",".f4p",".f4v",".fli",".flv",".mp4",
	".mov",".m4v",".mpg",".mpeg",".webm",".wmv",".mkv",".xvid"
	)
)
#endregion

#region Active Field Names
$FieldNames = 
@(
#Audio
	(
	"EFP_Directory","EFP_FileName","EFP_FileExtension","EFP_ItemType",
	"EFP_DateCreated","EFP_DateModified","EFP_DateAccessed","EFP_Title",
	"EFP_Album","EFP_TrackNo","EFP_RunTime","EFP_ContributingArtists",
	"EFP_YearReleased","EFP_Genre","EFP_Rating","EFP_BitRate",
	"EFP_Publisher","EFP_PartOfSet","EFP_Size"
	),
#Image
	(
	"EFP_Directory","EFP_FileName","EFP_FileExtension","EFP_ItemType",
	"EFP_DateCreated","EFP_DateModified","EFP_DateAccessed","EFP_DateTaken",
	"EFP_CameraMaker","EFP_CameraModel","EFP_Dimensions","EFP_BitDepth",
	"EFP_HorizontalResolution","EFP_VerticalResolution","EFP_Width ",
	"EFP_Height","EFP_EXIF_Version","EFP_ExposureBias","EFP_ExposureProgram",
	"EFP_ExposureTime","EFP_FStop","EFP_FlashMode","EFP_FocalLength",
	"EFP_FocalLength35mm","EFP_ISO_Speed","EFP_LightSource","EFP_MaxAperture",
	"EFP_MeteringMode","EFP_Orientation","EFP_ProgramMode","EFP_Saturation",
	"EFP_WhiteBalance","EFP_Size"
	), 
#Video
	(
	"EFP_Directory","EFP_FileName","EFP_FileExtension","EFP_ItemType",
	"EFP_DateCreated","EFP_DateModified","EFP_DateAccessed","EFP_Title",
	"EFP_Subtitle","EFP_Genre","EFP_RunTime","EFP_Album","EFP_ContributingArtists",
	"EFP_Publisher","EFP_YearReleased","EFP_ParentalRating","EFP_Rating","EFP_FrameWidth",
	"EFP_FrameHeight","EFP_FrameRate","EFP_VideoOrientation","EFP_VideoCompressionGUID",
	"EFP_DataRate","EFP_Bitrate","EFP_TotalBitrate","EFP_Size"
	)
)
#endregion

#region Parameter Arrays
$HashKeys =
@(
#Audio 
	(
	"FolderPath","Name","FileExtension","ItemType",
	"DateCreated","DateModified","DateAccessed","Title","Album",
	"TrackNo","RunTime","ContributingArtists","YearReleased",
	"Genre","Rating","BitRate","Publisher","PartOfSet","Size"
	),
#Image
	(
	"Folderpath","Name","FileExtension","ItemType",
	"DateCreated","DateModified","DateAccessed","DateTaken",
	"CameraMaker","CameraModel","Dimensions","BitDepth",
	"HorizontalResolution","VerticalResolution","Width","Height",
	"EXIF_Version","ExposureBias","ExposureProgram","ExposureTime",
	"FStop","FlashMode","FocalLength","FocalLength35mm",
	"ISO_Speed","LightSource","MaxAperture","MeteringMode",
	"Orientation","ProgramMode","Saturation","WhiteBalance","Size"
	),
#Video
	(
	"Folderpath","Name","FileExtension","ItemType",
	"DateCreated","DateModified","DateAccessed",
	"Title","Subtitle","Genre","Length","Album",
	"ContributingArtists","Publisher","Year",
	"ParentalRating","Rating","FrameWidth",
	"FrameHeight","FrameRate","VideoOrientation",
	"VideoCompressionGUID","DataRate","Bitrate",
	"TotalBitrate","Size"
	)
)
$PropNames =
@(
#Audio
	(
	"Folder path","Name","File extension","Item type",
	"Date created","Date modified","Date accessed",
	"Title","Album","#","Length","Contributing artists",
	"Year","Genre","Rating","Bit rate","Publisher",
	"Part of a compilation","Size"
	),
#Image
	(
	"Folder path","Name","File extension","Item type",
	"Date created","Date modified","Date accessed","Date taken",
	"Camera maker","Camera model","Dimensions","Bit depth",
	"Horizontal resolution","Vertical resolution","Width","Height",
	"EXIF version","Exposure bias","Exposure program","Exposure time",
	"F-stop","Flash mode","Focal length","35mm focal length","ISO speed",
	"Light source","Max aperture","Metering mode","Orientation",
	"Program mode","Saturation","White balance","Size"
	),
#Video
	(
	"Folder path","Name","File extension","Item type",
	"Date created","Date modified","Date accessed",
	"Title","Subtitle","Genre","Length","Album",
	"Contributing artists","Publisher","Year",
	"Parental rating","Rating","Frame width","Frame height",
	"Frame rate","Video orientation","Video compression",
	"Data rate","Bit rate","Total bitrate","Size"
	)
)
#endregion

#region Utility Functions
function Get-RegistryValue($key, $value) 
{(Get-ItemProperty -Path $key -Name $value).$value}

function Get-UserMediaFolder
{
	$SubKeys = @("My Music","My Pictures","My Video")
	$KeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
	Get-RegistryValue  $KeyPath $SubKeys[$ModeIdx]
}

function Get-Dir
{
	param(
		[Parameter(Mandatory=$True)][Alias('P')][String]$Path,
		[Parameter(Mandatory=$False)][Alias('F')][Switch]$File,
		[Parameter(Mandatory=$False)][Alias('R')][Bool]$Recurse=$False)

	$Flags = @(' ',' ')
	if($Recurse -eq $True){$Flags[0] = "-Recurse"}
	if($File.IsPresent){$Flags[1] = "-File"}
	$Cmd = "Get-ChildItem -Path {0} {1} {2}" -F $('"'+$Path+'"'),$Flags[0],$Flags[1]
	# Compile String $Cmd to a ScriptBlock & Invoke It --------
	Invoke-Command -ScriptBlock ($ExecutionContext.InvokeCommand.NewScriptBlock($Cmd))
	# ---------------------------------------------------------
}

function Get-ColumnHeaders
{
	$FormatStr = Get-FormatString
	$FormatStr -F $FieldNames[$ModeIdx]
}

function Get-FormatString
{
	for($C=0;$C -lt $FieldNames[$ModeIdx].Count-1;$C++)
		{$FormatStr += "{$C}`t"}
	$FormatStr += "{$($FieldNames[$ModeIdx].Count-1)}"
	return $FormatStr
}

function Get-ReportFile
{
	param(   
		[Parameter(Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Audio','Image','Video')]
		[String]$Mode)
	return -join(".\",$Mode,"Report.txt")
}

function Show-MsgBox($Msg)
{
	[System.Windows.Forms.MessageBox]::Show(
		$Msg,$ScriptID,
		[Windows.Forms.MessageBoxButtons]::Ok,
		[Windows.Forms.MessageBoxIcon]::Information)
}

function Get-MediaHashTable
{
	param(   
		[Parameter(Mandatory = $True)][Int]$ModeNo,
		[Parameter(Mandatory = $True)][Array]$MetaDataIndex)
    
	$Hash = @{}
	for($C=0;$C -lt $HashKeys[$ModeNo].Count;$C++)
		{$Hash += @{$HashKeys[$ModeNo][$C] = Get-IndexByMetadataName $MetaDataIndex $PropNames[$ModeNo][$C]}}
	$Hash
}
#endregion

#region File Property Record Retrieval Script Block
<#
This Script Block when Run as a Job on a seperate thread
reduces the memory requirements of the script as a whole.
A custom PSObject is returned with the requested file 
properties as defined by the input Hash Table.
#>
$GetFileProps = {
	param(
		[Parameter(Mandatory=$True)][Alias('T')][String]$TargetFile,
		[Parameter(Mandatory=$True)][Alias('H')][HashTable]$Hash,
		[Parameter(Mandatory=$True)][Alias('R')][Int]$RemainingFiles,
		[Parameter(Mandatory=$True)][Alias('C')][Int]$Filecount)

	$shell = New-Object -COMObject Shell.Application
	$folder = Split-Path -Path $TargetFile
	$file = Split-Path -Path $TargetFile -Leaf
	$shellfolder = $shell.Namespace($folder)
	$shellfile = $shellfolder.ParseName($file)
	$PercentComp = ((($Filecount - $RemainingFiles)/$Filecount)*100)
	$Activity = -join ('Collecting Properties - Files Remaining: ',
		$RemainingFiles,' of ',$Filecount,
		' - Percent Complete: ','({0:N3}%)' -f $PercentComp)
	Write-Progress -Activity $Activity -PercentComplete $PercentComp -Status $TargetFile
	$NewObj = New-Object -TypeName PSObject
	$hash.Keys | 
	ForEach-Object {
		$value = $shellfolder.GetDetailsOf($shellfile, $hash.$_)
		if ($value -as [Double]) {$value = [Double]$value}
		Add-Member -InputObject $NewObj `
			-TypeName NoteProperty `
			-NotePropertyName "$_" `
			-NotePropertyValue $Value -Force
	}
	if($NewObj.FolderPath.Length -eq 0)
		{$NewObj.FolderPath = $folder}
	if($NewObj.FileExtension.Length -eq 0)
		{$NewObj.FileExtension = [IO.Path]::GetExtension($File)}
	return $NewObj
}
#endregion

function Export-MediaFileDetails
{
if($Directory.Length -eq 0){$Directory = Get-UserMediaFolder}
#Output Field Names unless Appending to output
If($OutFile.Length -eq 0){$OutFile = Get-ReportFile -Mode $Mode}
if($Append -eq $False -and $RawOutput -eq $False){Get-ColumnHeaders| Out-File -FilePath $OutFile}
#Add File Info to List
$List = New-Object -TypeName System.Collections.Generic.List[System.Object]
Get-Dir -Path $Directory -Recurse $Recurse -File| Sort-Object -Property DirectoryName, Name |
	ForEach-Object{
		if($MediaFileTypes[$ModeIdx].Contains($_.Extension))
			{
			Write-Progress -Activity (-join ('Building File List ... ',$List.Count,' Files')) -Status " "
			$FI = New-Object -TypeName PSObject -Property @{FullName = $_.FullName; DirectoryName = $_.DirectoryName}
			$List.Add($FI)
			}        
	}|Out-Null

if($LogFileCount.IsPresent -and !$RawOutput.IsPresent){
	$T = "{0} - [{1} - {2}]" -F -join($Mode,' Files: ',$List.Count),$CurrDate,$CurrTime
	if($LogAppend.IsPresent)
		{$T|Out-File -FilePath $LogFile -Append}
	else
		{$T|Out-File -FilePath $LogFile}}

$TotalFiles = $List.Count
$FormatStr = Get-FormatString
$HoldPath = ''
while($List.Count -gt 0)
	{
	$Path = $List[0].DirectoryName.ToLower()
	if($Path -ne $HoldPath){
		$HoldPath = $Path
		$Hash = Get-MediaHashTable `
			-ModeNo $ModeIdx `
			-MetaDataIndex $(Get-MetadataIndex -P $Path)}
 
	$RV = Start-Job `
			-Name AddDtl `
			-ScriptBlock $GetFileProps `
			-ArgumentList $List[0].FullName,$Hash,$List.Count,$TotalFiles|`
		  Receive-Job -Wait -AutoRemoveJob

	if($RawOutput.IsPresent)
		{$RV}
	else
		{
		#Write Results to output file
		switch($ModeIdx)
			{
			0 #Audio
				{
				$FormatStr -f `
				$RV.FolderPath,
				$RV.Name,
				$RV.FileExtension,
				$RV.ItemType,
				$RV.DateCreated,
				$RV.DateModified,
				$RV.DateAccessed,
				$RV.Title,
				$RV.Album,
				$RV.TrackNo,
				$RV.RunTime,
				$RV.ContributingArtists,
				$RV.YearReleased,
				$RV.Genre,
				$RV.Rating,
				$RV.BitRate,
				$RV.Publisher,
				$RV.PartOfSet,
				$RV.Size | Out-File -FilePath $OutFile -Append
				}
			1 #Image
				{
				$FormatStr -f `
				$RV.Folderpath,
				$RV.Name,
				$RV.FileExtension,
				$RV.ItemType,
				$RV.DateCreated,
				$RV.DateModified,
				$RV.DateAccessed,
				$RV.DateTaken,
				$RV.CameraMaker,
				$RV.CameraModel,
				$RV.Dimensions,
				$RV.BitDepth,
				$RV.HorizontalResolution,
				$RV.VerticalResolution,
				$RV.Width,
				$RV.Height,
				$RV.EXIF_Version,
				$RV.ExposureBias,
				$RV.ExposureProgram,
				$RV.ExposureTime,
				$RV.FStop,
				$RV.FlashMode,
				$RV.FocalLength,
				$RV.FocalLength35mm,
				$RV.ISO_Speed,
				$RV.LightSource,
				$RV.MaxAperture,
				$RV.MeteringMode,
				$RV.Orientation,
				$RV.ProgramMode,
				$RV.Saturation,
				$RV.WhiteBalance,
				$RV.Size | Out-File -FilePath $OutFile -Append
				}
			2 #Video
				{
				$FormatStr -f `
				$RV.Folderpath,
				$RV.Name,
				$RV.FileExtension,
				$RV.ItemType,
				$RV.DateCreated,
				$RV.DateModified,
				$RV.DateAccessed,
				$RV.Title,
				$RV.Subtitle,
				$RV.Genre,
				$RV.Length,
				$RV.Album,
				$RV.ContributingArtists,
				$RV.Publisher,
				$RV.Year,
				$RV.ParentalRating,
				$RV.Rating,
				$RV.FrameWidth,
				$RV.FrameHeight,
				$RV.FrameRate,
				$RV.VideoOrientation,
				$RV.VideoCompressionGUID,
				$RV.Bitrate,
				$RV.DataRate,
				$RV.TotalBitrate,
				$RV.Size | Out-File -FilePath $OutFile -Append
				}
			}
		}
	$List.RemoveAt(0)
	} #End While Loop
if($Notify.IsPresent){Show-MsgBox "Export Complete!"}
}

#region Test Routines
$Targets = @('File','GridView','Host')
function New-TableEntry
{
	param(
		[Parameter(Mandatory = $False)][Alias('I')][Int]$Index=0,
		[Parameter(Mandatory = $False)][Alias('F')][String]$Field="",
		[Parameter(Mandatory = $False)][Alias('H')][String]$HashKey="",
		[Parameter(Mandatory = $False)][Alias('P')][String]$Property="",
		[Parameter(Mandatory = $False)][Alias('T')][String]$Type="")
	return New-Object -TypeName PSObject -Property @{Index = $Index;Type = $Type;FieldName = $Field;HashKey = $HashKey;Property=$Property}
}

function New-MediaTypeTableEntry
{
	param(
		[Parameter(Mandatory = $False)][Alias('I')][Int]$Index=0,
		[Parameter(Mandatory = $False)][Alias('E')][String]$Extension="",
		[Parameter(Mandatory = $False)][Alias('T')][String]$Type="")
	return New-Object -TypeName PSObject -Property @{Index = $Index;Type = $Type;Extension = $Extension}
}

function List-ParameterTables
<# Used to Sync tables #>
{
param(   
	[Parameter(Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('File','GridView','Host')]
	[String]$OutputTo)

$TargetIdx = [Array]::IndexOf($Targets,$OutputTo)
$Obj = @()

for($R=0;$R -lt $FieldNames.Count;$R++)
	{
	for($C=0;$C -lt $FieldNames[$R].Count;$C++)
		{
		$Obj += New-TableEntry `
			-Index    $C `
			-Type     $ValidModes[$R] `
			-Field    $($FieldNames[$R][$C]) `
			-HashKey  $($HashKeys[$R][$C]) `
			-Property $($PropNames[$R][$C])
		}
	}

switch($TargetIdx)
	{
	0 #File
		{
		$Obj|Format-Table -AutoSize -Property Type,Index,FieldName,Hashkey,Property|
		Out-File -Filepath .\MetaDataTable.txt
		}
	1 #GridView
		{
		$Obj|Out-GridView -Title "Metadata Tables"
		}
	2 #Host
		{
		$Obj|Format-Table -AutoSize -Property Type,Index,FieldName,Hashkey,Property
		}
	}
}

function List-FileTypes
{
param(   
	[Parameter(Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('File','GridView','Host')]
	[String]$OutputTo)

$TargetIdx = [Array]::IndexOf($Targets,$OutputTo)
$Obj = @()

for($R=0;$R -le $MediaFileTypes.GetUpperBound(0);$R++)
	{
	for($C=0;$C -lt $MediaFileTypes[$R].Count;$C++)
		{
		$Obj += New-MediaTypeTableEntry `
			-Index    	$C `
			-Type     	$ValidModes[$R] `
			-Extension	$($MediaFileTypes[$R][$C]) `
		}
	}

switch($TargetIdx)
	{
	0 #File
		{
		$Obj|Format-Table -AutoSize -Property Type,Index,Extension|
		Out-File -Filepath .\FileTypeTable.txt
		}
	1 #GridView
		{
		$Obj|Out-GridView -Title "File Extension Tables"
		}
	2 #Host
		{
		$Obj|Format-Table -AutoSize -Property Type,Index,Extension
		}
	}
}
#endregion Test Routines

#Call Main function
switch($Debug){
	0{Export-MediaFileDetails}
	1{List-ParameterTables -OutputTo Host}
	2{List-FileTypes -OutputTo Host}
}
<# Test Commands
Measure-Command {.\Export-MediaFileDetailsEx -Mode Video -R}
Measure-Command {.\Export-MediaFileDetailsEx -Mode Video -R -RO |Out-GridView}
Measure-Command {.\Export-MediaFileDetailsEx.ps1 -Mode Video -Dir F:\Video -R -LC}
Measure-Command {.\Export-MediaFileDetailsEx -Mode Video -R -Dir "\\MYBOOKLIVE\Public\Shared Videos\Movies\"}
Measure-Command {.\Export-MediaFileDetailsEx -Mode Video -R -LC -Dir "\\MYBOOKLIVE\Public\Shared Videos\"}
#>