<#
.NOTES
Name:	UtilitiesLib.ps1
Author:  Randy Turner
Version: 1.0
Date:	12/04/2021

.SYNOPSIS
Provides a wrapper for utility fumctions used in multiple scripts 
#>

<#
.NOTES
Name:	 Convert-ArrayToString()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to get the TabIndex Info for a Control (Form) & all the child controls.
It's called recursively for each child where the HasChildren property is True.
Returns an ArrayList of objects with each controls: 
Name, Parent (Name), TabIndex, TabStop, & HasChildren

.PARAMETER Array
Required, TopMost control (Entry Point) of recursive call

.PARAMETER Delimiter
Optional, delimiter to place between Array elements in the resulting string.

.EXAMPLE
$FileTypes = Convert-ArrayToString -Array $AFileTypes
This example returns a comma delimited string.
#>
function Convert-ArrayToString{
	param(
		[Parameter(Mandatory)][Array]$Array,
		[Parameter()][String]$Delimiter=',')
	$FormatStr = ''
	for($C=0;$C -le $Array.GetUpperBound(0);$C++){$FormatStr += "{$C}$Delimiter"}
	return -join($FormatStr.SubString(0,$FormatStr.length - $Delimiter.length) -f $Array)
}

<#
.NOTES
Name:	 Get-ControlTabInfo()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2022
Revision History:
1.0 - 08/15/2022 - Initial Release

.SYNOPSIS
function used to get the TabIndex Info for a Control (Form) & all the child controls.
It's called recursively for each child where the HasChildren property is True.
Returns an ArrayList of objects with each controls: 
Name, Parent (Name), TabIndex, TabStop, & HasChildren

.PARAMETER TopControl Alias: T
Required, TopMost control (Entry Point) of recursive call

.EXAMPLE
Get-ControlTabInfo -TopControl $Form1|Sort-Object -Property Parent,TabIndex|FT
This example returns a table of objects.
#>
function Get-ControlTabInfo{
	param([Parameter(Mandatory)][Alias('T')][System.Windows.Forms.Control]$TopControl)
	$ArrayList = New-Object -TypeName Collections.ArrayList
	foreach($Control in $TopControl.Controls){
		$Entry = [PSCustomObject][Ordered]@{
			Name = $Control.Name;
			Parent = $Control.Parent.Name;
			TabIndex = $Control.TabIndex;
			TabStop = $Control.TabStop;
			HasChildren = $Control.HasChildren}
		[void]$ArrayList.Add($Entry)
		if($Control.HasChildren){Get-ControlTabInfo -T $Control}
    }
    $ArrayList
}

<#
.NOTES
Name:	 Get-EnumMembers()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2022
Revision History:
1.0 - 08/15/2022 - Initial Release

.SYNOPSIS
function used to get the members of an Enum their Name, Value, and the HostEnum (Name).
Returns an unfiltered set of PSCustomObjects the HostEnum is only on the first object 
returned to provide a breakpoint between sets.

.PARAMETER Enum Alias: E
Required, Enum to retrieve, passed as a type. May be passed via a pipe.

.EXAMPLE
PS> Get-EnumMembers -Enum ([Drawing.KnownColor])
This example returns a table of objects for the Windows.Drawing.KnownColor Enum.
Note that to pass Types they must be contained in some manner to be distinguished.

.EXAMPLE
ps> $MyEnums|Get-EnumMembers
This example returns a table of objects for multiple Enums.
#>
function Get-EnumMembers{
	param([Parameter(Mandatory,ValueFromPipeline)][Alias('E')][Type]$Enum)
	Begin{$PrevName = $Null}
	Process{
		[Enum]::GetNames($Enum.FullName) | ForEach-Object{
			$Cmd = '[{0}]::{1}.Value__' -f $Enum.FullName,$_
			$Value = Invoke-Command -ScriptBlock ($ExecutionContext.InvokeCommand.NewScriptBlock($Cmd))
			$Obj = [PSCustomObject][Ordered]@{Name = $_;Value = $Value}
			if($PrevName -ne $Enum.FullName){
				Add-Member -InputObject $Obj -MemberType NoteProperty -Name HostEnum -Value $Enum.FullName
				$PrevName = $Enum.FullName}
			$Obj
		}
	}
}

<#
.NOTES
Name:	 Get-RegistryValue()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to get a value from a registry key.

.PARAMETER Key 
Required, Host registry key.

.PARAMETER Value
Required, Registry key node to retrieve value.
.EXAMPLE
PS> Get-RegistryValue -Key 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Value 'My Video'
This example returns a path to the Current User's 'My Video' folder.
#>
function Get-RegistryValue{
	param(
		[Parameter(Mandatory)][String]$Key,
		[Parameter(Mandatory)][String]$Value)
	(Get-ItemProperty -Path $Key).$Value
}


<#
.NOTES
Name:	 Get-SelectedFontStyle()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to get a value from a registry key.

.PARAMETER SelectedFontStyle Alias: FS 
Required, Selected Font Style to retrieve.


.EXAMPLE
PS> Get-SelectedFontStyle -FS 'BoldItalic'
This example returns a 'BoldItalic' Font Style.
#>
function Get-SelectedFontStyle{
	param([Parameter(Mandatory)][Alias('FS')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Bold','Italic','BoldItalic','Regular')]
		[String]$SelectedFontStyle)

	$MyParam = (Get-Command -Name $MyInvocation.MyCommand).Parameters
	$Styles = $MyParam['SelectedFontStyle'].Attributes.ValidValues
	$DFS = [Drawing.FontStyle]

	switch([Array]::IndexOf($Styles,$SelectedFontStyle))
	{
		0	{$DFS::Bold}
		1	{$DFS::Italic}
		2	{($DFS::Bold -bor $DFS::Italic)}
		Default {$DFS::Regular}
	}
}


<#
.NOTES
Name:	 Get-SplitFilename()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to split a path into Filename or Extension only.

.PARAMETER FileName Alias: Fn
Required, Input file path.

.PARAMETER ExtensionOnly Alias: Eo
Optional, switch causes the file extension only to return

.EXAMPLE
PS> Get-SplitFilename -FN $path -EO 
This example returns a file extension for the input path.
#>
function Get-SplitFilename{
	param(
		[Parameter(Mandatory)][Alias('Fn')][String]$FileName,
		[Parameter()][Alias('Eo')][Switch]$ExtensionOnly)
	return $(if($ExtensionOnly -eq $False)
			{[IO.Path]::GetFileNameWithoutExtension($FileName)}
		else
			{[IO.Path]::GetExtension($FileName)})
}

<#
.NOTES
Name:	 Load-Dll()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to split a path into Filename or Extension only.

.PARAMETER FileName Alias: DLL
Required, Input file path.

.PARAMETER FullPath Alias: Full
Optional, switch indicates if FileName is a Fully Qualified path.

.PARAMETER ShowErrorMsg Alias: Msg
Optional, switch causes an error message to display on occurence

.EXAMPLE
PS> Get-SplitFilename -FN $path -EO 
This example returns a file extension for the input path.
#>
function Load-Dll{
	param(
		[Parameter(Mandatory)][Alias('DLL')][String]$FileName,
		[Parameter()][Alias('Full')][Switch]$FullPath,
		[Parameter()][Alias('Msg')][Switch]$ShowErrorMsg)

	$DLLPath = if($FullPath.IsPresent){$Filename}
		else{Resolve-CurrentLocation $Filename}
	
	if($ShowErrorMsg.IsPresent -and [IO.File]::Exists($DLLPath) -eq $False){
		$ErrMsg = "DLL File: {0} Missing or Invalid - Job Aborted!" -f $DLLPath
		Show-MessageBox -M $ErrMsg -T $Form1.Text -I Stop
		exit
	}
	Add-Type -Path $DLLPath
}


<#
.NOTES
Name:	 Resolve-CurrentLocation()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to resolve the fullpath of a .\ relative path.

.PARAMETER Path Alias: P
Required, .\ relative path to resolve

.EXAMPLE
PS> Resolve-CurrentLocation -P .\Images
This example returns a Fully Qualified path to the Directory\File
#>
function Resolve-CurrentLocation{
	param([Parameter(Mandatory)][Alias('P')][String]$Path)
	return "{0}\{1}" -f (Get-Location),(Split-Path -Path $($Path) -Leaf)
}

<#
.NOTES
Name:	 Split-EnumNames()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2022
Revision History:
1.0 - 08/15/2022 - Initial Release

.SYNOPSIS
function used to get an array of strings each string is an Enum Name which has been expanded
based upon "Pascal Casing" with a space inserted between words. Thus ResetColumnWidths becomes
'Reset Column Widths'. this allows the use of Enum member names as text wherever needed, for
example as a button label or menuitem label.


.PARAMETER Enum Alias: E
Required, Enum to retrieve, passed as a type

.EXAMPLE
$FileMenuText = Split-EnumNames -Enum ([FileMenuItem])
This example returns an array of strings for the emun.
Note that to pass Types they must be contained in some manner to be distinguished.
#>
function Split-EnumNames{
	param([Parameter(Mandatory)][Alias('E')][Type]$Enum)
	[Enum]::GetNames($Enum.FullName)|%{
		$Parts = [Regex]::Matches($_,'[A-Z][a-z]+').Value
		$Mask = ''
		for($C = 0;$C -lt $Parts.Count; $C++){$Mask += ('{'+$('{0}' -f $C)+'} ')}
		$Mask = $Mask.Substring(0,$Mask.Length - 1)
		$RV = $Mask -f $Parts
		$RV
	}
}

<#
.NOTES
Name:	 Toggle-Boolean()
Author:  Randy Turner
Version: 1.0
Date:	08/15/2010
Revision History:
1.0 - 08/15/2010 - Initial Release

.SYNOPSIS
function used to resolve the fullpath of a .\ relative path.

.PARAMETER Target
Required, Boolean to toggle value.

.EXAMPLE
PS> $Boolin = Toggle-Boolean -Target $Boolin
This example returns the inverse value of the input boolean.
#>
function Toggle-Boolean{
	param([Parameter(Mandatory)][Bool]$Target)
	return ($Target = !$Target)
}