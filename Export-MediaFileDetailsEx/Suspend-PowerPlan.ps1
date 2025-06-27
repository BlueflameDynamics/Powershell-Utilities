<#
.Notes
 ----------------------------------------------------------------------------------
 Script: Suspend-PowerPlan.ps1
 Version: 2.0
 Author: Randy Turner
 Date: 06/02/2018
 Description: Helper Function to Suspend Power Plan when running PowerShell scripts
 Comments:
 ----------------------------------------------------------------------------------
 
.Synopsis
	Function to suspend your current Power Plan settings when running a PowerShell script.

.DESCRIPTION

	Function to suspend your current Power Plan settings when running a PowerShell script.
	the Display & System Switches may be used without Continuous to temporarily reset the
	timeout counter(s) to zero (Must be repeated periodically) or Disabled until reset when
	paired with Continuous. Awaymode will always set the Continuous flag regardless.
	Continuous when specified alone or in the absense of any flags will reset your
	Power Plan to normal.

.PARAMETER AwayMode Alias: A
	Enables AwayMode implies Continuous, if omitted. Continuous Required by AwayMode.

.PARAMETER Display Alias: D
	Reset Display Timeout, Locked when paired with Continuous

.PARAMETER System Alias: S
	Reset System Timeout, Locked when paired with Continuous

.PARAMETER Continuous Alias: C
	Continuous when paired with the other switches locks them until reset by Continuous alone.
	Assumed in the absense of any Switches (Reset).

.LINK
	https://msdn.microsoft.com/en-us/library/windows/desktop/aa373208(v=vs.85).aspx

.EXAMPLE
	Suspend-PowerPlan -Continuous -System
	Prevent system from sleeping

.EXAMPLE
	Suspend-PowerPlan -System -Continuous -Display
	Prevent system from sleeping and keep display alive

.EXAMPLE
	Suspend-PowerPlan -System -Continuous -AwayMode
	set system in AwayMode (-Continuous implied if not specified, Required by AwayMode)

.EXAMPLE
	Suspend-PowerPlan
	Clear all flags on the current thread defaults to -Continuous
#>
function Suspend-PowerPlan
{
	param
		(
		[Parameter()][Alias('A')][switch]$Away,
		[Parameter()][Alias('D')][switch]$Display,
		[Parameter()][Alias('S')][switch]$System,
		[Parameter()][Alias('C')][switch]$Continuous
		)

$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
  public static extern void SetThreadExecutionState(uint esFlags);
'@

	$Settings = @(0,0,0,0) #All flags off
	$W32 = Add-Type -MemberDefinition $code -Name System -Namespace Win32 -PassThru 
	$ES_Continuous = [uint32]"0x80000000" 
	<#Requests that the other EXECUTION_STATE flags set remain in effect until
	SetThreadExecutionState is called again with the ES_CONTINUOUS flag set and
	one or more of the other EXECUTION_STATE flags cleared.#>
	$ES_AwayMode_Required = [uint32]"0x00000040" #Requests Away Mode to be enabled. Must be used with ES_Continuous flag.
	$ES_Display_Required  = [uint32]"0x00000002" #Requests display availability (display idle timeout is prevented).
	$ES_System_Required   = [uint32]"0x00000001" #Requests system availability (sleep idle timeout is prevented).

	if($Away.IsPresent)   {$Settings[1] = $ES_AwayMode_Required}
	if($Display.IsPresent){$Settings[2] = $ES_Display_Required}
	if($System.IsPresent) {$Settings[3] = $ES_System_Required}
	if($Continuous.IsPresent -or $Away.IsPresent -or (($Settings[1] -bor $Settings[2] -bor $Settings[3]) -eq 0))
		{$Settings[0] = $ES_Continuous}
	$W32::SetThreadExecutionState($Settings[0] -bor $Settings[1] -bor $Settings[2] -bor $Settings[3])  
}