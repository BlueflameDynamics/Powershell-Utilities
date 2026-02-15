Function Get-OSVersion{
	$L = Build-WinList(7,8,10,11)
	$C = Get-CimInstance Win32_Operatingsystem|Select-Object -Expand Caption
	$V = $L[$C.Substring(10,10)]
	$D = [Ordered]@{Caption = $C}
	$L.Values|%{$D.Add(('IsWin'+$_),$False)}
	$D.Add('Version',$V)
	$D['IsWin'+$V] = $True
	[PSCustomObject]$D
}
Function Build-WinList([Int[]]$V){
	$L = [Ordered]@{}
	ForEach($I in $V){$L.Add(("Windows {0}" -f $I).PadRight(10).Substring(0,10), $I)}
	$L
}