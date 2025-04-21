Enum IconCatalogGroup{
	Tag
	ControlIndex
}
Class IconCatalogItem{
#Object Properties
[String]$Tag = ''
[Int]$ControlIndex = 0
[Int]$IconIndex = 0

#Explicitly declared New method
IconCatalogItem(){}

#Overload New method
IconCatalogItem([String]$Tag,[Int]$ControlIndex,[Int]$IconIndex){
	$This.Tag = $Tag
	$This.ControlIndex = $ControlIndex
	$This.IconIndex = $IconIndex}

#Convert 2D Array to an array of IconCatalogItem objects
[IconCatalogItem[]] Static ConvertFrom2dArray([Array]$Items){
	return @(for($C = 0;$C -lt $Items[0].Count;$C++){
		[IconCatalogItem]::New(
			$Items[[IconCatalogGroup]::Tag][$C],
			$Items[[IconCatalogGroup]::ControlIndex][$C],
			$C)})}

#Convert an OrderedDictionary to an array of IconCatalogItem objects
[IconCatalogItem[]] Static ConvertFromOrderedDictionary([System.Collections.Specialized.OrderedDictionary]$OrderedDictionary){
	$C = 0
	return @(ForEach($Entry in $OrderedDictionary.GetEnumerator()){
		[IconCatalogItem]::New(
			$Entry.Key,
			$Entry.Value,
			$C++)})}

#Filter IconCatalogItem object Array and Sort
[IconCatalogItem[]] Static GetSelectedItems([IconCatalogItem[]]$Items){
	return ($Items |
		Where-Object -Property ControlIndex -NotMatch -Value -1 |
		Sort-Object -Property ControlIndex)}
}
Function Test-IconCatalogItem{
Param([Parameter()][ValidateNotNullOrEmpty()][ValidateSet('Array','OrderedDictionary')][String]$Mode = 'Array')

#region Function ValidateSet Access
$MyParam=(Get-Command -Name $MyInvocation.MyCommand).Parameters
$Modes=$MyParam['Mode'].Attributes.ValidValues
$ModeIdx=[Array]::IndexOf($Modes,$Mode)
#endregion

Switch($ModeIdx){
	0	{
		$IconCatalog = @(
		(
		'App Icon','Open Playlist','New Playlist','Edit Playlist','Reload Playlist','Directory',
		'Audio File','Find','Find Next','Font Settings','Help','Exit','Info','Play','Pause','Stop','Logo'),
		(-1,0,1,2,3,-1,-1,4,5,6,-1,8,7,-1,-1,-1,-1)
		)
		$IconCatalog|Out-Host
		('{0}IconCatalog[0] Bounds: {1} Type: {2}' -f "`r`n",$IconCatalog[0].GetUpperBound(0),$IconCatalog[0][0].GetType())|Out-Host
		('{0}IconCatalog[1] Bounds: {1} Type: {2}' -f "`r`n",$IconCatalog[1].GetUpperBound(0),$IconCatalog[1][0].GetType())|Out-Host
		$IconCatalogItems = [IconCatalogItem]::ConvertFrom2dArray($IconCatalog)
		}
	1	{
		$IconCatalog = [Ordered]@{
		'App Icon' = -1;'Open Playlist' = 0 ;'New Playlist' = 1;'Edit Playlist' = 2;
		'Reload Playlist' = 3;'Directory' = -1;'Audio File' = -1;'Find' = 4;
		'Find Next' = 5;'Font Settings' = 6;'Help' = -1;'Exit' = 8;
		'Info' = 7;'Play' = -1;'Pause' = -1;'Stop' = -1;'Logo' = -1}
		$IconCatalog|Out-Host
		$IconCatalogItems = [IconCatalogItem]::ConvertFromOrderedDictionary($IconCatalog)
		}
	}

$IconCatalogItems|Out-Host
$IconCatalogItems = [IconCatalogItem]::GetSelectedItems($IconCatalogItems)
$IconCatalogItems|Out-Host
}