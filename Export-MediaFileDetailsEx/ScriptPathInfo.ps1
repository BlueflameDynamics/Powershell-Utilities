
function Get-ScriptPathInfo
{
	param([Parameter(Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("FullPath","Directory","File","FileName","Extension","*")]
		[String]$Mode)

	$Modes = @("FullPath","Directory","File","FileName","Extension","*")
	$ModeIdx = [array]::IndexOf($Modes,$Mode)

	if($MyInvocation.ScriptName.Length -gt 0) {
	    $RV=$MyInvocation.ScriptName
	    $NewObj = New-Object PSObject
	    Add-Member -InputObject $NewObj -MemberType NoteProperty -Name FullPath		-Value $($RV)
	    Add-Member -InputObject $NewObj -MemberType NoteProperty -Name Directory	-Value $([System.IO.Path]::GetDirectoryName($RV))
	    Add-Member -InputObject $NewObj -MemberType NoteProperty -Name File			-Value $([System.IO.Path]::GetFileName($RV))
	    Add-Member -InputObject $NewObj -MemberType NoteProperty -Name FileName		-Value $([System.IO.Path]::GetFileNameWithoutExtension($RV))
	    Add-Member -InputObject $NewObj -MemberType NoteProperty -Name Extension	-Value $([System.IO.Path]::GetExtension($RV))

	    Switch($ModeIdx)
		    {
		    0 {Return $NewObj.FullPath}
		    1 {Return $NewObj.Directory}
		    2 {Return $NewObj.File}
		    3 {Return $NewObj.FileName}
		    4 {Return $NewObj.Extension}
		    5 {Return $NewObj} 
		    }
    }
}

function Get-ScriptPathInfoTest
{
$SN = @(0..5)
$SN[0] = Get-ScriptPathInfo -Mode FullPath
$SN[1] = Get-ScriptPathInfo -Mode Directory
$SN[2] = Get-ScriptPathInfo -Mode File
$SN[3] = Get-ScriptPathInfo -Mode FileName
$SN[4] = Get-ScriptPathInfo -Mode Extension
$SN[5] = Get-ScriptPathInfo -Mode *
$SN
}