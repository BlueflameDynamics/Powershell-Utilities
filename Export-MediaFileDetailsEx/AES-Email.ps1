<#
.NOTES
	File Name:	AES-Email.ps1
	Version:	1.2 - 05/20/2022
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	05/21/2011
	History:
		V1.2 - 05/20/2022
		V1.1 - 05/21/2021	
		V1.0 - 05/21/2011

.SYNOPSIS
	<SynopsisText>

.DESCRIPTION
	This Script contains functions for:
	* Creating an AES encryption key file
	* Creating an encrypted Email password file
	* Retrieving Email credentials used to send an email
	* Sending an SMTP email with or without an attachment
	
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
----------------------------------------------------------------------------------------
#>

#Settings Configuration Object
$Settings = [PSCustomObject][Ordered]@{
	KeyFile = '<Enter Path to Key File Here>'
	PasswordFile = '<Enter Path to Password File Here>'
	SmtpPort = 587 #SMTP Port Number
	SmtpServer = '<Enter SMTP Email Server Here>'
	TargetEmail = '<Enter Target Email Address Here>'
	SentBy = '<Enter Sending Email Address Here>'
}

<#
.NOTES
	Function:	Create-AESKeyFile
	Version:	1.0 - 05/21/2011
	Author:		Randy Turner

.SYNOPSIS
	Creates an AES KeyFile

.PARAMETER $KeyLength Alias: L
	Optional, Key Length of 16(Default),24, or 32 Bytes.

.INPUTS
	Preconfigured Settings Object

.OUTPUTS
	AES Key File	
#>
function Create-AESKeyFile{
	Param([Parameter()][Alias('L')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet(16,24,32)]
		[Int]$KeyLength=16)
# Create AES key with random data and export to file
$Key = New-Object Byte[] $KeyLength
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
$Key | Out-File $Settings.KeyFile
}

<#
.NOTES
	Function:	Create-EmailPwdFile
	Version:	1.0 - 05/21/2011
	Author:		Randy Turner

.SYNOPSIS
	Creates an AES Encripted Email Password File

.PARAMETER SmtpPassword Alias: SmtpPwd
	Required, Smtp Password for sending account.

.INPUTS
	Preconfigured Settings Object


.OUTPUTS
	AES Encripted Email Password File	
#>
function Create-EmailPwdFile{
Param([Parameter(Mandatory)][Alias('SmtpPwd')][String]$SmtpPassword)
# Create SecureString object
$Key = Get-Content $Settings.KeyFile
$Password = $SmtpPassword | ConvertTo-SecureString -AsPlainText -Force
$Password | ConvertFrom-SecureString -key $Key | Out-File $Settings.PasswordFile
}

<#
.NOTES
	Function:	Get-EmailCredentials
	Version:	1.0 - 05/21/2011
	Author:		Randy Turner
	
.SYNOPSIS
	Gets the Decripted Email Credentials for sending email

.INPUTS
	Preconfigured Settings Object
#>
function Get-EmailCredentials{
# Create PSCredential object
$key = Get-Content $Settings.KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential `
 -ArgumentList $Settings.SentBy, (Get-Content $Settings.PasswordFile | ConvertTo-SecureString -Key $key)
Return $MyCredential
}

<#
.NOTES
	Function:	Send-Email3
	Version:	1.0 - 05/21/2011
	Author:		Randy Turner

.SYNOPSIS
	Sends an email message from sender to the Target Email Address.

.PARAMETER FromAddr Alias: From
	Optional, Sending Email Address

.PARAMETER ToAddr Alias: To
	Optional, Receiving Email Address

.PARAMETER Subject Alias: Sub
	Required, Email Subject

.PARAMETER Body Alias: Msg
	Required, Email Message Body

.PARAMETER Attachment Alias: Att
	Optional, Email Attachment

.INPUTS
	Preconfigured Settings Object	
#>
function Send-Email3{
Param(
	[Parameter()][Alias('From')][String]$FromAddr=$Settings.SentBy,
	[Parameter()][Alias('To')][String]$ToAddr=$Settings.TargetEmail,
	[Parameter(Mandatory)][Alias('Sub')][String]$Subject,
	[Parameter(Mandatory)][Alias('Msg')][String]$Body,
	[Parameter()][Alias('Att')][String]$Attachment)
	$Credentials = Get-EmailCredentials
	if($Attachment.Length -eq 0)
		{
		Send-MailMessage -From $FromAddr -To $ToAddr -Subject $Subject `
		-Body $Body -SmtpServer $Settings.SmtpServer -Port $Settings.SmtpPort -UseSsl `
		-Credential ($Credentials)
		}
	else
		{
		Send-MailMessage -From $FromAddr -To $ToAddr -Subject $Subject `
		-Body $Body -SmtpServer $Settings.SmtpServer -Port $Settings.SmtpPort -UseSsl `
		-Credential ($Credentials) -Attachments $Attachment
		}
	}