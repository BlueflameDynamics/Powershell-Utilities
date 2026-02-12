#Sample invocation: & {.\AppRegistryTypeAccelerators.ps1}
# CTA = CustomTypeAccelerator
# Functions in this module use CTA as the canonical abbreviation.

# ==================================
# Custom TypeAccelerator Initializer
# ==================================

# Sentinel: tracks whether initialization has already occurred
if (-not $script:CTA){

	# Acquire the internal TypeAccelerators class
	# The use of reflection is required 
	$script:CTA = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')

	# Register accelerators only once per session
	$script:CTA::Add('W32Reg',     [Microsoft.Win32.Registry])
	$script:CTA::Add('W32RegKey',  [Microsoft.Win32.RegistryKey])
	$script:CTA::Add('W32RegKind', [Microsoft.Win32.RegistryValueKind])
}

# Utility: refresh the live accelerator dictionary
function Update-CTAList {
	$script:CTAList = $script:CTA::Get
}

# Utility: test for a specific accelerator
function Test-CTAExists([string]$Key) {
	$script:CTA::Get.ContainsKey($Key)
}

# Utility: test whether the loader has already run
function Test-CTALoaded {
	[bool]$script:CTA
}

function Add-CTA([string]$Key,[type]$Value){
	if ([string]::IsNullOrWhiteSpace($Key)){
		throw 'Invalid Key, must not be null or White Space.'
	}
	if ($script:CTA::Get.ContainsKey($Key)){
		throw ('TypeAccelerator [{0}] already exists.' -f $Key)
	}
	$script:CTA::Add($Key,$Value)	
}