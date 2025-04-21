param(
		[Parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet("All","Audio","Image","Video")]
			[String]$Mode,
		[Parameter(Mandatory=$False)][Alias('R')][Switch]$Reboot)

	$Modes = @("All","Audio","Image","Video")
	$ModeIdx = [Array]::IndexOf($Modes,$Mode)
	$EmailTxt = @(
		"Get-NetMediaInfo Completed!",
		"Get-NetMediaInfo of TPSI-NET Completed!",
		".\EventTimeLog.txt")
	$Locations = @(
		"\\MYBOOKLIVE\Public\Shared Music\",
		"\\MYBOOKLIVE\Public\Shared Pictures\",
		"\\MYBOOKLIVE\Public\Shared Videos\")

function Suspend-PowerPlan
{
	param
		(
		[Parameter(Mandatory=$false)][Alias('A')][switch]$Away,
		[Parameter(Mandatory=$false)][Alias('D')][switch]$Display,
		[Parameter(Mandatory=$false)][Alias('S')][switch]$System
		)

	$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
  public static extern void SetThreadExecutionState(uint esFlags);
'@

	$settings = @(0,0,0)
	$ste = Add-Type -memberDefinition $code -name System -namespace Win32 -passThru 
	$ES_CONTINUOUS = [uint32]"0x80000000" #Requests that the other EXECUTION_STATE flags set remain in effect until SetThreadExecutionState is called again with the ES_CONTINUOUS flag set and one of the other EXECUTION_STATE flags cleared.
	$ES_AWAYMODE_REQUIRED = [uint32]"0x00000040" #Requests Away Mode to be enabled.
	$ES_DISPLAY_REQUIRED = [uint32]"0x00000002" #Requests display availability (display idle timeout is prevented).
	$ES_SYSTEM_REQUIRED = [uint32]"0x00000001" #Requests system availability (sleep idle timeout is prevented).

	if($Away.IsPresent)   {$settings[0] = $ES_AWAYMODE_REQUIRED}
	if($Display.IsPresent){$settings[1] = $ES_DISPLAY_REQUIRED}
	if($System.IsPresent) {$settings[2] = $ES_SYSTEM_REQUIRED}
	$ste::SetThreadExecutionState($ES_CONTINUOUS -bor $settings[0] -bor $settings[1] -bor $settings[2])  
}    

function Play-Sound
{
param(
    [Parameter(Mandatory=$True)][Alias('F')][String]$SoundFile,
    [Parameter(Mandatory=$False)][Alias('D')][int]$Delay=3)

$SoundLib = "h:\Software\Sounds\"
(new-object Media.SoundPlayer "$SoundLib\$SoundFile").play();
Start-Sleep -Seconds $Delay
}

function Get-Audio
	{.\Export-MediaFileDetailsEx -Mode Audio -Dir $Locations[0] -Log $EmailTxt[2] -R -LC -LA}
 
function Get-Image
	{.\Export-MediaFileDetailsEx -Mode Image -Dir $Locations[1] -Log $EmailTxt[2] -R -LC -LA}

function Get-Video
	{.\Export-MediaFileDetailsEx -Mode Video -Dir $Locations[2] -Log $EmailTxt[2] -R -LC -LA}

function Get-All
	{
	$ShowStatus = {
		param([String]$Mode,[Int]$ModeIdx)
		Clear-Host
		"Step ($Mode) $ModeIdx of 3 Complete!"
		#Flush Memory Variables
		Get-Variable -Exclude PWD,*Preference|Remove-Variable -EA 0
		}
	Get-Audio; Invoke-Command -Scriptblock $ShowStatus -ArgumentList $Modes[1],1
	Get-Image; Invoke-Command -Scriptblock $ShowStatus -ArgumentList $Modes[2],2
	Get-Video; Invoke-Command -Scriptblock $ShowStatus -ArgumentList $Modes[3],3
	}

function Get-NetMediaInfo
	{
	Import-Module  -Name .\AES_Email.ps1 -Force
	$TM=""
    Play-Sound -SoundFile "\AUTHORIZ.WAV" -Delay 5
    Suspend-PowerPlan -System
	Switch($ModeIdx)
		{
		0 {$TM=Measure-Command -Expression {Get-All}}
		1 {$TM=Measure-Command -Expression {Get-Audio}}
		2 {$TM=Measure-Command -Expression {Get-Image}}
		3 {$TM=Measure-Command -Expression {Get-Video}}
		}
	$TM|Out-File -FilePath $($EmailTxt[2]) -Append
	"Sending Conformation Email, Please wait ..."| Out-Host
	Send-Email3 -Subject $EmailTxt[0] -Body $EmailTxt[1] -Att $EmailTxt[2]
	"Transmission Complete ..."| Out-Host
	Remove-Item  -Path $($EmailTxt[2])
    Suspend-PowerPlan #Reset
    Play-Sound -SoundFile "\ANALYSIS.WAV" -Delay 5
	if($Reboot.IsPresent){Start-Sleep -Seconds 30;Restart-Computer}
	}

	#Execute Main Function
	Get-NetMediaInfo