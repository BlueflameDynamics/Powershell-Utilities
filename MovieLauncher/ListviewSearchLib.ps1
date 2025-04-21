<#
.NOTES
Name:	 New-LvwSearchValueItem Function
Author:  Randy Turner
Version: 1.0
Date:	 08/15/2022

.SYNOPSIS
This class is a custom object used to store values used in searching a Listview control.
The object has 4 properties: SeekValue, Index, Column, and Initialized. 
SeekValue holds the value to seek. Index is the integer value of the Listview.Items index
representing the next item or row to search. Column is the Listview.subitems index or column to search. 
Lastly Initialized is a boolean used to indicate if the properties of the LvwSearchValueItem have been
initialized by a previous search and is awaiting a Find Next request.
#>
Class LvwSearchValueItem{
	[String]$SeekValue = ''
	[Int]$Index = 0
	[Int]$Column = 0
	[Boolean]$Initialized = $False
}

<#
.NOTES
Name:	 Find-ListViewItem Function
Author:  Randy Turner
Version: 1.1
Date:	 12/01/2007

.SYNOPSIS
This function executes a search operation of a WinForms Listview control. This search is
of the Listview.Items.Text unless the Listview.View is set to Details. If the View is in
Details mode the search is of the Listview as a 2 dimensionl table where the 
SearchValueItem.Index is the Row and LvwSearchValueItem.Column is the column to search.
During a FindNext operation the FindStr,StartPos, & Column parameters are taken from
the associated LvwSearchValueItem properties.

.PARAMETER SearchValueItem Alias: SVI
Required, Custom LvwSearchValueItem used to control the search.
Must be passed by reference [Ref].

.PARAMETER LvwToSearch Alias: Lvw
Required, is a Listview control to search.

.PARAMETER FindStr Alias: Val
Required for Starting Find, is the string value to find within the Listview.

.PARAMETER StartPos Alias: Row
Optional, is the Listview.Items Index representing the
starting location (Row) of the search operation.

.PARAMETER Column Alias: Col
Optional, is the Listview.SubItems Index representing the
Column within a 2-Dimessional table search when the Listview.View is set to Details.
#>
function Find-ListViewItem{
	param(
		[Parameter(Mandatory)][Alias('SVI')][Ref]$SearchValueItem,
		[Parameter(Mandatory)][Alias('Lvw')][Windows.Forms.ListView]$LvwToSearch,
		[Parameter()][Alias('Val')][String]$FindStr='',
		[Parameter()][Alias('Row')][Int]$StartPos=0,
		[Parameter()][Alias('Col')][Int]$Column=0)

	if($SearchValueItem.Value.Initialized -eq $True){
		$StartPos = $SearchValueItem.Value.Index
		$FindStr = $SearchValueItem.Value.SeekValue
		$Column = $SearchValueItem.Value.Column
	}
	elseif($FindStr -eq ''){
		Write-Error 'Value to Seek, FindStr Parameter Missing & Required!' -ErrorAction Stop
	}
	
	for ($R = $StartPos; $R -lt $LvwToSearch.Items.Count; $R++){
		[Windows.Forms.ListViewItem]$Item=$LvwToSearch.Items[$R]
		if(($LvwToSearch.View -ne [Windows.Forms.View]::Details -and 
			$Item.Text.ToLower().Contains($FindStr.Tolower())) -or 
			($LvwToSearch.View -eq [Windows.Forms.View]::Details -and 
			$Item.SubItems[$Column].Text.ToLower().Contains($FindStr.Tolower()))){
				$Item.Selected = $True
				$Item.EnsureVisible()
				$SearchValueItem.Value.Index = $R + 1
				$SearchValueItem.Value.Column = $Column
				$SearchValueItem.Value.SeekValue = $FindStr
				$SearchValueItem.Value.Initialized = $True
				$R = $LvwToSearch.Items.Count}
		elseif($R -eq $LvwToSearch.Items.Count - 1){
			$SearchValueItem.Value = [LvwSearchValueItem]::New()
			Show-MessageBox -M $("[{0}] Not Found!" -f $FindStr) -T $Form1.Text -I Information}
	}
}