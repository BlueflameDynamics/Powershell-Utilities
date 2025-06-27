function Test-InternetConnection{
	param([Parameter()][Alias('T')][String]$Target='WWW.Novell.com') 
	$result = Test-Connection -ComputerName $Target -ErrorAction SilentlyContinue
	if($null -eq $result){$False}else{$True}
}
