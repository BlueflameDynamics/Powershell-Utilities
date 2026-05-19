<#
.NOTES
	File Name:	Set-TVTitles.ps1
	Version:	1.5b - 11/17/2023
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	02/03/2023

.SYNOPSIS
	This Script will list the TV meta tags for m4v & mp4 files,
	Set the 'Title' tag in the format: <TVShow>: S<Season number>E<Episode number> - <Episode Title>,
	rename the input files to include the 'Episode Title'.

.DESCRIPTION
	This Script will list the TV meta tags for m4v & mp4 files,
	Set the 'Title' tag in the format: <TVShow>: S<Season number>E<Episode number> - <Episode Title>,
	rename the input files to include the 'Episode Title' ie: <TVShow>_S<Season number>E<Episode number>_<Episode Title>.<Ext>.
	The rename is actually a copy to a new location preserving the original files. The 'Title' tag is based upon
	the corresponding tags.
	
	The MP4 file format has 3 basic file extensions: MP4, M4V, & M4A.
	
	In practice an MP4 file may contain both Audio and Video or one and not the other. 
	
	By MPEG standards an MP4 file may not contain any DRM. Apple wanted to use the MP4 file format 
	but use their FairPlay-DRM and thus Apple created the M4A & M4V extensions as alternatives to
	the MP4 extension to further distiguish between files with Video and Audio or Video only
	and files with Audio only (the M4A extension). By using an extension of m4v or m4a Apple
	couldn't be accused of violating MPEG standards.
	
	As far as the internal file structure, nothing is different. Both M4A & M4V are 
	MP4 files. In practice the distinction of Non-DRM and DRM between the extensions has been
	lost. Some players will work with all 3 extensions, others will insist on a specific
	extension. Likewise some players will only show chapter information for m4v files.
	
	This script was originally designed to process only m4v files because it was designed to process
	Video files. Generally speaking you may simply rename an MP4 to use any of the 3
	extensions you require.

.COMPONENT
	Minimum Required Components: 
	TagLibSharp.dll (v2.1.0.0), BlueflameDynamics.VideoTags.dll (v1.0.0),
	Microsoft.WindowsAPICodePack.dll (v1.1.5), Microsoft.WindowsAPICodePack.Shell.dll (v1.1.5)

.PARAMETER MediaDirectory Alias: Dir
	Required, This is the directory containing the input files to process

.PARAMETER Mode
	Mode of operation: TVShows (Default) or Movies

.PARAMETER SeasonWidth Alias: SW
	This is the minimun number of digits in the TV season number 

.PARAMETER EpisodeWidth Alias: EW
	This is the minimun number of digits in the TV episode number

.PARAMETER ApplyFix Alias: A
	This switch will cause the file 'Title' Tag to be updated with a string
	in the format:  <TV Show Name>: S<Season Number>E<Episode Number> - <Episode Title>

.PARAMETER RenameFiles Alias: R
	This switch will cause the input files to be copied to a 'New' directory in the same
	location as the input media files and rename the files in the format:  
	<TV Show Name>_S<Season No.>E<Episode No.>_<Episode Title>.<Ext>. Any embedded spaces
	are replaced by an '_' as are any chains of '...', all illegal filename characters
	and any single quotes are removed from the new filename. When used may be coupled with
	the Dynamic Parameters: 'SuppressProgressWindow' alias: 'SP', a switch which will cause
	the output files to be copied without displaying Windows' Progress Window, and 
	'TargetDirectory' alias: 'TD', a path which overrides the default target location for
	the file copy operation.

.INPUTS
	Video files with a supported extension: m4v or mp4

.OUTPUTS
	Varies based upon script parameters passed	

.EXAMPLE
	PS> Set-TVTitles.ps1 -Dir <Input Directory> 
	Returns an array of PSCustomObjects listing TV related metadata tags

.EXAMPLE
	PS> Set-TVTitles.ps1 -Dir <Input Directory> -Mode Movies
	Returns an array of PSCustomObjects listing Movie related metadata tags
	
.EXAMPLE
	PS> Set-TVTitles.ps1 -Dir <Input Directory> -Mode TVShows -ApplyFix
	Enables update of the Title Meta Tag for TV

.EXAMPLE
	PS> Set-TVTitles.ps1 -Dir <Input Directory> -RenameFiles 
	Enables copying the files to a 'New' directory renaming the files.
#>
[CmdletBinding()]
Param(
		[Parameter(Mandatory)][Alias('Dir')][String]$MediaDirectory,
		[Parameter()]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('TVShows','Movies')]
			[String]$Mode='TVShows',
		[Parameter()][Alias('SW')][Int]$SeasonWidth=1,
		[Parameter()][Alias('EW')][Int]$EpisodeWidth=2,
		[Parameter()][Alias('A')][Switch]$ApplyFix,
		[Parameter()][Alias('R')][Switch]$RenameFiles)
DynamicParam{
	# Set up the Run-Time Parameter Dictionary
	$RuntimeParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::New()

	if($RenameFiles.IsPresent){
		#region SuppressProgressWindow 
		$DynamicParamName = 'SuppressProgressWindow'
		$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::New()
		$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::New()
		$ParameterAttribute.Mandatory = $false
		$AttributeCollection.Add($ParameterAttribute)
		$ParameterAlias = [System.Management.Automation.AliasAttribute]::New('SP')
		$AttributeCollection.Add($ParameterAlias)
		$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::New($DynamicParamName, [Switch], $AttributeCollection) 
		$RuntimeParameterDictionary.Add($DynamicParamName, $RuntimeParameter)
		#endregion
		#region TargetDirectory
		$DynamicParamName = 'TargetDirectory'
		$AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::New()
		$ParameterAttribute = [System.Management.Automation.ParameterAttribute]::New()
		$ParameterAttribute.Mandatory = $false
		$ParameterAttribute.Position = 6 
		$AttributeCollection.Add($ParameterAttribute)
		$ParameterAlias = [System.Management.Automation.AliasAttribute]::New('TD')
		$AttributeCollection.Add($ParameterAlias)
		$RuntimeParameter = [System.Management.Automation.RuntimeDefinedParameter]::New($DynamicParamName, [String], $AttributeCollection)
		$RuntimeParameter.Value = Join-Path -Path $MediaDirectory -ChildPath '\New\'
		$RuntimeParameterDictionary.Add($DynamicParamName, $RuntimeParameter)
		#endregion
	}
	# When done building dynamic parameters, return dictionary
	return $RuntimeParameterDictionary
}

Begin {
	#region Module Imports		
	Import-Module -Name .\Exists.ps1 -Force
	Import-Module -Name .\Invoke-CopyFile.ps1 -Force
	Import-Module -Name .\MetadataIndexLib.ps1 -Force
	Import-Module -Name .\TagLib-Sharp.dll -Force <#Note: Requires V2.1.0 as V2.3.0 Currupts files.#>
	Import-Module -Name .\BlueflameDynamics.VideoTags.dll -Force
	#endregion

	#region Custom Enums
	Enum OpMode{
		TVShows
		Movies
	}
	Enum MediaKinds{ #ToDo: check itunes for PodCast & Voice Memo
		Not_Set = -1 #stik is Null
		Home_Movie = 0
		Music = 1
		AudioBook = 2
		Music_Video = 6
		Movie = 9
		TV_Show = 10
		Booklet = 11
		Ringtone = 14
	}
	#endregion

	#region Apple BoxType Constants
	$BOXTYPE_TVSH = 'tvsh' # TV Show or Series name
	$BOXTYPE_TVSN = 'tvsn' # season number
	$BOXTYPE_TVEN = 'tven' # episode name
	$BOXTYPE_TVES = 'tves' # episode number
	$BOXTYPE_TVNN = 'tvnn' # TV Network
	$BOXTYPE_PURD = 'purd' # purchase date
	$BOXTYPE_DESC = 'desc' # short description
	$BOXTYPE_LDES = 'ldes' # long description
	$BOXTYPE_STIK = 'stik' # iTunes Media Kind Index (MediaKinds Enum)

	$BOXTYPE_BROADCASTDATE = [TagLib.ByteVector]::New()
	$BOXTYPE_BROADCASTDATE.Add([TagLib.ByteVector]::FromUShort(169)[1])
	$BOXTYPE_BROADCASTDATE.Add([TagLib.ByteVector]::FromString('d'))
	$BOXTYPE_BROADCASTDATE.Add([TagLib.ByteVector]::FromString('a'))
	$BOXTYPE_BROADCASTDATE.Add([TagLib.ByteVector]::FromString('y'))

	$BOXTYPE_EPISODE_TITLE = [TagLib.ByteVector]::New()
	$BOXTYPE_EPISODE_TITLE.Add([TagLib.ByteVector]::FromUShort(169)[1])
	$BOXTYPE_EPISODE_TITLE.Add([TagLib.ByteVector]::FromString('n'))
	$BOXTYPE_EPISODE_TITLE.Add([TagLib.ByteVector]::FromString('a'))
	$BOXTYPE_EPISODE_TITLE.Add([TagLib.ByteVector]::FromString('m'))
	#endregion

	#region Extended File Properties
	$SysTags = @(
	'Title','Subtitle','Genre','Length','Contributing artists',
	'Parental rating','Rating','Frame width','Frame height',
	'Frame rate','Bit rate','Data rate','Size'
	)
	#endregion

	#region Utility Functions
	Function Remove-InvalidFileNameChars{
		Param([Parameter(Mandatory)][String]$Str)
		$InvalidFileNameChars = [System.IO.Path]::GetInvalidFileNameChars()
		ForEach($Char in $InvalidFileNameChars){
			$Str = $Str.Replace($Char.ToString(), '')}
		Return $Str
	}

	Function Get-ExtendedFileProperties{
		Param([Parameter(Mandatory)][Alias('P')][String]$Path)

		$ShellJob = {
			Param(
				[Parameter(Mandatory)][Alias('P')][String]$Folder,
				[Parameter(Mandatory)][Alias('F')][String]$File,
				[Parameter(Mandatory)][Alias('I')][Int[]]$PropIndices)
			#Create Windows Shell Object
			$Shell = New-Object -COMObject Shell.Application
			$ShellFolder = $Shell.Namespace($Folder)
			$ShellFile = $ShellFolder.ParseName($File)
			$RV = @()
			foreach($Index in $PropIndices)
				{$RV+=$ShellFolder.GetDetailsOf($ShellFile, $Index)}
			return $RV}

		[String[]]$PropNames = $SysTags
		# Get Property Index Numbers
		$MetaDataIndex = Get-MetadataIndex $Path
		[Int[]]$PropIndex = @()
		for($C = 0;$C -le $PropNames.GetUpperBound(0);$C++)
			{$PropIndex += Get-IndexByMetadataName -MetaDataIndex $MetaDataIndex -SearchValue $PropNames[$C]}
		#Parse Path
		$Folder = Split-Path -Path $Path
		$File = Split-Path -Path $Path -Leaf
		$RV = Start-Job `
				-Name 'Get-EFP' `
				-ScriptBlock $ShellJob `
				-ArgumentList $Folder,$File,$PropIndex| `
			  Receive-Job -Wait -AutoRemoveJob
		return $RV
	}

	Function Get-VideoTags{
		param([Parameter(Mandatory)][Alias('P')][String]$Path)
		$RV = [Ordered]@{}
		$C = 0
		$Values = Get-ExtendedFileProperties -Path $Path
		ForEach($Tag In $SysTags){
			$RV[$Tag] += $Values[$C++]}
		Return $RV
	}

	Function Get-MediaType{
		$MediaKind = -1 # Not Set
		$Array = $AppleTag.DataBoxes($BOXTYPE_STIK)
		if (($Array -ne $Null) -and ($Array.Data.Count -ge 1)){
		    $MediaKind = $Array.Data[0]}
		$MediaKindStr = [MediaKinds].GetEnumName($MediaKind).Replace('_',' ')
		Return (('ID: {0} - Name: {1}' -f $MediaKind,$MediaKindStr))
	}

	Function Get-NumericDataboxValue{
		Param([Parameter(Mandatory)][Alias('BX')][String]$BoxType)
		$Number = $null
		$Array = $AppleTag.DataBoxes($BoxType)
		if(($Array -ne $Null) -and ($Array.Data.Count -ge 3)){
			$Number = [Int]$Array.Data[3]}
		Return $Number
	}
	#endregion

	#region Get Script Parameter ValidValues
	$MyParam=($MyInvocation.MyCommand).Parameters
	$Modes=$MyParam['Mode'].Attributes.ValidValues
	$ModeIdx=[Array]::IndexOf($Modes,$Mode)
	#endregion
}

Process{
	$Output = @()
	$ContentRatingOrg = -join ($(if($ModeIdx -eq [OpMode]::Movies){'MPAA'}else{'US-TV'}),' Rating')
	$FileList = Get-ChildItem -Path ($MediaDirectory+'\*') -Include @('*.m4v','*.mp4')|Sort-Object -Property Fullname
	$AirDate = [Ordered]@{Name = -join ($(if($ModeIdx -eq [OpMode]::Movies){'Release'}else{'Air'}),' Date');Date = $Null}

	if($RenameFiles.IsPresent){
		$TargetDirectory = $RuntimeParameterDictionary['TargetDirectory'].Value.ToString()
		if(!$TargetDirectory.EndsWith('\')){$TargetDirectory += '\'}
	}

	ForEach ($File in $FileList){
		$VideoTags = Get-VideoTags -Path $File.FullName
		$MP4Tags = [BlueflameDynamics.VideoTags]::GetVideoTags($File.FullName)
		$Mediafile = [TagLib.File]::Create($File.FullName)
		[TagLib.Mpeg4.AppleTag]$AppleTag = $MediaFile.GetTag([TagLib.TagTypes]::Apple, 1)
		$RowData = [PSCustomObject][Ordered]@{Directory = $File.Directory;File = $File.Name}

		if($RenameFiles.IsPresent -and $ModeIdx -eq [OpMode]::TVShows){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'New Filename' -Value $Null}

		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Media Type' -Value (Get-MediaType)
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Images' -Value $Mediafile.Tag.Pictures.Count
		
		if($ModeIdx -eq [OpMode]::TVShows){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'TV Show' -Value ($AppleTag.DataBoxes($BOXTYPE_TVSH).Text)}
			
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Title' -Value ($AppleTag.DataBoxes($BOXTYPE_EPISODE_TITLE).Text)
		if($VideoTags['Subtitle'].Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Subtitle' -Value $VideoTags['Subtitle']}

		if((Get-Module -ListAvailable -Name Get-MediaInfo) -ne $Null){
			$Studio = (Get-MediaInfoValue -Path $File.FullName -Kind General -Index 0 -Parameter ProductionStudio)
			$CR = (Get-MediaInfoValue -Path $File.FullName -Kind General -Index 0 -Parameter ContentRating)
			$CR = $CR.ToUpper().Split('|')
			$ContentRatingOrg = '{0} Rating' -f $CR[0]
			if($Studio.Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Studio' -Value $Studio}}

		if($ModeIdx -eq [OpMode]::TVShows){		
			if($AppleTag.DataBoxes($BOXTYPE_TVNN).Text.Length -gt 0){
				Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Network' -Value ($AppleTag.DataBoxes($BOXTYPE_TVNN).Text)}}

		$AirDate.Date = $AppleTag.DataBoxes($BOXTYPE_BROADCASTDATE).Text
		if(($AirDate.Date -ne $Null) -and ($AirDate.Date.Length -gt 10)){
			$AirDate.Date = $AirDate.Date.SubString(0,10)
			}

		if($VideoTags['Parental Rating'].Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name $ContentRatingOrg -Value $VideoTags['Parental Rating']}

		if($AirDate.Date.Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name $AirDate.Name -Value $AirDate.Date}

		if($ModeIdx -eq [OpMode]::TVShows){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Episode Name' `
				-Value ($AppleTag.DataBoxes($BOXTYPE_TVEN).Text)
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Season' `
				-Value (Get-NumericDataboxValue -BX $BOXTYPE_TVSN)
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name '#' `
				-Value (Get-NumericDataboxValue -BX $BOXTYPE_TVES)
		}

		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Run Time' -Value $VideoTags['Length']
		if($Mediafile.Tag.Artists.Count -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Artists' -Value $Mediafile.Tag.Artists}
		if($Mediafile.Tag.Genres.Count -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Genre' -Value $Mediafile.Tag.Genres}
		if($VideoTags['Rating'] -ne 'Unrated'){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Rating' -Value $VideoTags['Rating']}
		$Frame = [System.Drawing.Size]::New($VideoTags['Frame Width'],$VideoTags['Frame Height'])
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Frame Size' -Value $Frame
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Frame Rate' -Value $VideoTags['Frame Rate']
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Audio Channels' -Value $MP4Tags.Audio.ChannelCount
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Audio Sample Rate' -Value ('{0}Khz' -f $MP4Tags.Audio.SampleRate)
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Video Data Rate' -Value $VideoTags['Data rate']
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Audio Bit Rate' -Value $VideoTags['Bit rate']
		Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Size' -Value $VideoTags['Size']

		if($AppleTag.DataBoxes($BOXTYPE_DESC).Text.Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Short Description' -Value ($AppleTag.DataBoxes($BOXTYPE_DESC).Text)}
		if($AppleTag.DataBoxes($BOXTYPE_LDES).Text.Length -gt 0){
			Add-Member -InputObject $RowData -MemberType NoteProperty -Name 'Long Description' -Value ($AppleTag.DataBoxes($BOXTYPE_LDES).Text)}

		if(($RenameFiles.IsPresent -and $ModeIdx -eq [OpMode]::TVShows) -and $RowData.'Episode Name'.Length -gt 0){
			if(!(Test-Exists -Mode Directory -Location $TargetDirectory)){
				[Void][System.IO.Directory]::CreateDirectory($TargetDirectory)}
			$EN = $RowData.'Episode Name'.Replace(' ','_')
			$EN = Remove-InvalidFileNameChars ($EN.Replace("'",''))
			$EN = $EN.Replace('...','_')
			$NewName = '{0}{1}_{2}{3}' -f `
				$TargetDirectory,
				[System.IO.Path]::GetFileNameWithoutExtension($File.Name),
				$EN,
				$File.Extension
			$RowData.'New Filename' = $NewName
			Write-Host -Object ('Copying: {0} to: {1}' -f $File.Name,$NewName)
			if($RuntimeParameterDictionary['SuppressProgressWindow'].IsSet){
				Copy-Item -LiteralPath $File.FullName -Destination $NewName}
			else{
				Invoke-CopyFile -Source $File.FullName -Target $NewName}
			$Mediafile.Dispose()
			$Mediafile = [TagLib.File]::Create($NewName)
		}

		if($ApplyFix.IsPresent -and $ModeIdx -eq [OpMode]::TVShows){
			if($RowData.'TV Show'.Contains(':')){$Token = ' - '}else{$Token = ': '}
			$Fmt = -join ('{0}',$Token,'S{1:D',$SeasonWidth,'}E{2:D',$EpisodeWidth,'} - {3}')
			$RowData.Title = ($Fmt -f $RowData.'TV Show',$RowData.Season,$RowData.'#',$RowData.'Episode Name')
			$Mediafile.Tag.Title = $RowData.Title
			$Mediafile.Tag.Title
			$Mediafile.Save()
			$Mediafile.Dispose()
			Remove-Variable -Name Mediafile
			}
		$Output += $RowData
	}
}

End{if(($ModeIdx -eq [OpMode]::TVShows -and $ApplyFix -eq $False) -or $ModeIdx -eq [OpMode]::Movies){$Output}}

<# Test Commands
.\Set-TVTitles.ps1 -MediaDirectory '\\MyCloud1\Public\Shared Videos\Movies\Alien_Nation\' -Mode Movies
.\Set-TVTitles.ps1 -MediaDirectory 'f:\Video\Space_1999\' -Mode TVShows -ApplyFix
.\Set-TVTitles.ps1 -MediaDirectory 'f:\Video\Space_1999\' -Mode TVShows | Out-GridView -Title 'Set-TVTitles Report'
#>