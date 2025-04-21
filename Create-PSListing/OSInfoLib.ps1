Function Get-OSVersion{
	$HT = Build-HashTable(7,8,10,11)
	$S = Get-CimInstance Win32_Operatingsystem|Select-Object -Expand Caption
	$V = $HT[$S.Substring(10,10)]
	$T = [Ordered]@{Caption = $S}
	$HT.Keys.Replace(' ','')|%{$T.Add(('IsWin'+$_.Substring(7)),$False)}
	$T.Add('Version',$V)
	$T['IsWin'+$V] = $True
	[PSCustomObject]$T
}
Function Build-HashTable([Int[]]$V){
	$T = [Ordered]@{}
	ForEach($I in $V){$T.Add(('Windows {0} ' -f $I).Substring(0,10),$I)}
	$T
}