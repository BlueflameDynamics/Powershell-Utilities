<#
.NOTES
Function Name: Get-SplitPath
Author:  Randy Turner
Version: 1.0
Date:	08/04/2007

.SYNOPSIS
This Function is designed to split a Directory\File path into it's component parts,
which are returned as a [PSCustomObject] & it's properties:
FullPath,Directory,File(Name with extension),FileName(without extension), & Extension.
Unlike the C\C++ splitpath() the Drive component isn't broken out in order for the
function to support the use of UNC, URL, & URI paths.

.DESCRIPTION
This Function is designed to split a Directory\File path into it's component parts.

.PARAMETER Path
This parameter is the path to split. Pipeline parameter passing is supported.

.EXAMPLE
Get-SplitPath -Path $InputFilePath
This example will return an object whose properties are the path parts.
#>
function Get-SplitPath{
	Param([Parameter(Mandatory,ValueFromPipeline)][String]$Path)
	[PSCustomObject][Ordered]@{
		FullPath=[IO.Path]::GetFullPath($Path)
		Directory=[IO.Path]::GetDirectoryName($Path)
		File=[IO.Path]::GetFileName($Path)
		FileName=[System.IO.Path]::GetFileNameWithoutExtension($Path)
		Extension=[IO.Path]::GetExtension($Path)}
}