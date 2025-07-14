<#
.NOTES
	File Name:	Get-WindowsProductKeyEx.ps1
	Version:	2.0 - 2025/07/12
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	2021/05/20

.SYNOPSIS
	This script gets basic information about the currently installed Windows version and
	the Windows Update History from the Registry. Two of the reported properties are:
	'ProductID' (SerialNumber) & 'ProductKey'.

.DESCRIPTION
	This script gets basic information about the currently installed Windows version and
	the Windows Update History from the Registry. Note: Windows Update does not
	always update: ProductName & ReleaseId in Registry. Where the Registry based values may
	be vwrong the Common Information Model (CIM) property values are reported.
	The Update History is shown as a table. This scripy includes functions for viewing the
	CIM properties and the Windows NT CurrentVersion Registry properties for those intrested.
	
	----------------------------------------------------------------------------------------
	Security Note: This is an unsigned script, Powershell security may require you run the
	Unblock-File cmdlet with the Fully qualified filename before you can run this script,
	assuming PowerShell security is set to RemoteSigned.
	---------------------------------------------------------------------------------------- 
#> 

#region Common Variables
$RegKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$Win32OS = 'Win32_OperatingSystem'
#endregion

#region Diagnostic Helper Functions
function Show-WindowsOSInformation{
	$FileOut = $Env:Temp + '\OS_Info.htm' 
	$OSObj = Get-CimInstance -class $Win32OS
	$OSObj.CimSystemProperties | ConvertTo-Html  -Title 'Windows OS Information' | Out-File $FileOut
	$OSObj.CimInstanceProperties | ConvertTo-Html | Out-File $FileOut -Append
	Invoke-item $FileOut
	Start-Sleep -Seconds 5 
	Remove-Item $FileOut
}

function Get-Win32_OperatingSystemProperties{
	$OSObj = Get-CimInstance -class $Win32OS
	$OSObj.CimInstanceProperties|Sort-Object -Property Name
}

function Get-WindowsCurrentVersion{
	(Get-ItemProperty $RegKey)
}
#endregion

#region Define Helper Function to Decode Product Key
function Get-WindowsProductKey{
	#Test whether this is Windows 7 or older:
	function Test-Win7{
	  $OSVersion = [System.Environment]::OSVersion.Version
	  ($OSVersion.Major -eq 6 -and $OSVersion.Minor -lt 2) -or
	  $OSVersion.Major -le 6
	}

	#Declare c# decoder 
	$code = @'
		// original implementation: https://github.com/mrpeardotnet/WinProdKeyFinder
		using System;
		using System.Collections;

		public static class Decoder
		{
			  public static string DecodeProductKeyWin7(byte[] digitalProductId)
			  {
				  const int keyStartIndex = 52;
				  const int keyEndIndex = keyStartIndex + 15;
				  var digits = new[]
				  {
					  'B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'M', 'P', 'Q', 'R',
					  'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9',
				  };
				  const int decodeLength = 29;
				  const int decodeStringLength = 15;
				  var decodedChars = new char[decodeLength];
				  var hexPid = new ArrayList();
				  for (var i = keyStartIndex; i <= keyEndIndex; i++)
				  {
					  hexPid.Add(digitalProductId[i]);
				  }
				  for (var i = decodeLength - 1; i >= 0; i--)
				  {
					  // Every sixth char is a separator.
					  if ((i + 1) % 6 == 0)
					  {
						  decodedChars[i] = '-';
					  }
					  else
					  {
						  // Do the actual decoding.
						  var digitMapIndex = 0;
						  for (var j = decodeStringLength - 1; j >= 0; j--)
						  {
							  var byteValue = (digitMapIndex << 8) | (byte)hexPid[j];
							  hexPid[j] = (byte)(byteValue / 24);
							  digitMapIndex = byteValue % 24;
							  decodedChars[i] = digits[digitMapIndex];
						  }
					  }
				  }
				  return new string(decodedChars);
			  }

			  public static string DecodeProductKey(byte[] digitalProductId)
			  {
				  var key = String.Empty;
				  const int keyOffset = 52;
				  var isWin8 = (byte)((digitalProductId[66] / 6) & 1);
				  digitalProductId[66] = (byte)((digitalProductId[66] & 0xf7) | (isWin8 & 2) * 4);

				  const string digits = "BCDFGHJKMPQRTVWXY2346789";
				  var last = 0;
				  for (var i = 24; i >= 0; i--)
				  {
					  var current = 0;
					  for (var j = 14; j >= 0; j--)
					  {
						  current = current*256;
						  current = digitalProductId[j + keyOffset] + current;
						  digitalProductId[j + keyOffset] = (byte)(current/24);
						  current = current%24;
						  last = current;
					  }
					  key = digits[current] + key;
				  }

				  var keypart1 = key.Substring(1, last);
				  var keypart2 = key.Substring(last + 1, key.Length - (last + 1));
				  key = keypart1 + "N" + keypart2;

				  for (var i = 5; i < key.Length; i += 6)
				  {
					  key = key.Insert(i, "-");
				  }

				  return key;
			  }
		 }
'@
	#Compile c#:
	Add-Type -TypeDefinition $code

	#Get raw product key:
	$digitalId = (Get-ItemProperty -Path $RegKey -Name DigitalProductId).DigitalProductId
 
	#Use required static c# method:
	if(Test-Win7)
		{[Decoder]::DecodeProductKeyWin7($digitalId)}
	else
		{[Decoder]::DecodeProductKey($digitalId)}
}
#endregion

function Get-OSUpdateHistory{
	#Get old versions
	(Get-ChildItem -Path HKLM:\System\Setup\Source* | ForEach-Object {Get-ItemProperty -Path Registry::$_}) +
	#Append current version
	(Get-ItemProperty $RegKey) |
	#Select the fields we care about
	Select-Object ProductName, ReleaseId, CurrentBuild,UBR,DisplayVersion, 
	#Make date readable and convert from UTC
	@{n='InstallDate'; e={([DateTime]'1/1/1970').AddSeconds($_.InstallDate).ToLocalTime()}} |
	#Sort
	Sort-Object -Property InstallDate
}

#region MainLine
#Build a custom System Information Object
$SystemInfo = [PSCustomObject][Ordered]@{
	'ComputerName' = (Get-CimInstance $Win32OS).CSName;
	'Description' = (Get-CimInstance $Win32OS).Description;
	'ProductName' = (Get-CimInstance $Win32OS).Caption;
	'EditionID' = (Get-ItemProperty $RegKey).EditionID
	'CurrentVersion' = '{0}.{1}.{2}.{3}' -f (
		(Get-ItemProperty $RegKey).CurrentMajorVersionNumber,
		(Get-ItemProperty $RegKey).CurrentMinorVersionNumber,
		(Get-ItemProperty $RegKey).CurrentBuild,
		(Get-ItemProperty $RegKey).UBR); #Update Build Revision
	'FeatureID' = (Get-ItemProperty $RegKey).DisplayVersion; 
	'InstallDate' = (Get-CimInstance $Win32OS).InstallDate;
	'Architecture' = (Get-CimInstance $Win32OS).OSArchitecture;
	'ProductID' = (Get-ItemProperty $RegKey).ProductID;
	'ProductKey' = Get-WindowsProductKey
}
$SystemInfo
Get-OSUpdateHistory|Format-Table -AutoSize
'Note: Windows Update does not always update: ProductName & ReleaseId in Update History.'
#endregion