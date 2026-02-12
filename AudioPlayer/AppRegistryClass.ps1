& {.\AppRegistryTypeAccelerators.ps1}

#region Enums
enum ConstKey{
	DefaultCompanyName
	DefaultAppName
	RegistryRoot
}
enum MsgID{
	ApplicationKeyNonExistant
	MustBeNonEmptyString
	MustBeRegistryValueKind
}
enum RegistryHives{
	CurrentUser
	LocalMachine
}
#endregion

Class AppRegistryCfg{
	[string]$DefaultCompanyName = '<Your_Company_Name_Here>'
	[string]$DefaultAppName     = '<Your_App_Name_Here>'
	[string]$RegistryRoot       = 'Software'
	[string]$CompanyName        = $this.DefaultCompanyName
	[bool]$CompanyLock          = $false

	AppRegistryCfg(){}
	AppRegistryCfg([string]$CompanyName){
		if([string]::IsNullOrWhiteSpace($CompanyName)){
			throw [ArgumentException]::new('CompanyName must be a valid non-empty string.')
		}
		$this.CompanyName = $CompanyName
		$this.CompanyLock = $true
	}
}

Class AppRegistry{
#region Fields
	hidden [AppRegistryCfg]$Cfg = [AppRegistryCfg]::new()
	hidden [string]$Hive        = [AppRegistry]::RegistryHive([RegistryHives]::CurrentUser)
	hidden [bool]$WriteMode     = $false
	[string]$CompanyName        = $this.Cfg.CompanyName
	[string]$AppName            = $this.Cfg.DefaultAppName
	[W32RegKey]$ApplicationKey = $null
	[W32RegKey]$CompanyKey     = $null
#endregion

#region Constructors
	# Default - HKCU, read-only.
	AppRegistry(){}

	# HKCU, Write selectable
	AppRegistry([bool]$Writeable){
		$this.WriteMode = $Writeable
	}

	# HKCU/HKLM, read-only
	AppRegistry([RegistryHives]$Hive){
		$this.Hive = [AppRegistry]::RegistryHive($Hive)
	}

	# HKCU/HKLM, Write selectable
	AppRegistry([RegistryHives]$Hive,[bool]$Writeable){
		$this.Hive      = [AppRegistry]::RegistryHive($Hive)
		$this.WriteMode = $Writeable
	}

	# HKCU/HKLM, CompanyName, Write selectable
	AppRegistry([string]$CompanyName,[RegistryHives]$RegHive,[bool]$Writeable){
		if([string]::IsNullOrWhiteSpace($CompanyName)){
			throw [ArgumentException]::new(
				[AppRegistry]::MsgStore([MsgID]::MustBeNonEmptyString) -f 'CompanyName'
			)
		}

		# Instance-local configuration
		$this.Cfg         = [AppRegistryCfg]::new($CompanyName)
		$this.Hive        = [AppRegistry]::RegistryHive($RegHive)
		$this.WriteMode   = $Writeable
		$this.CompanyName = $this.Cfg.CompanyName
		$this.AppName     = $this.Cfg.DefaultAppName
	}
#endregion

#region Class Literals
static [string]$EmptyString  = [string]::Empty
static [string]$DefaultValue = [string]::Empty
hidden [W32RegKey]$HiveSoftwareRoot = $null
#endregion

#region Catalogs
	static hidden [string]RegistryHive([RegistryHives]$Key){
		$Hives = [Enum]::GetNames([RegistryHives])
		return $Hives[$Key]
	}

	static hidden [string]MsgStore([MsgID]$Key){
		$Map = @(
			'Application Key {0}:{1} is Non-Existant',
			'{0} must be a valid non-empty string.',
			'Value must be a {0} for RegistryValueKind.{1}.'
		)
		return $Map[$Key]
	}
#endregion

#region Helpers
	[string]GetCompanyRootPath(){
		return '{0}\{1}' -f `
			$this.Cfg.RegistryRoot,
			$this.CompanyName
	}

	[string]GetAppRootPath(){
		return '{0}\{1}\{2}' -f `
			$this.Cfg.RegistryRoot,
			$this.CompanyName,
			$this.AppName
	}

	static hidden [string]GetHiveAbbrev([string]$Key){
		$Abbrevs = @{
			'CurrentUser'  = 'HKCU'
			'LocalMachine' = 'HKLM'
		}
		return $Abbrevs[$Key]
	}

	hidden [bool]IsDefaultCompany(){
		return $this.CompanyName -eq $this.Cfg.DefaultCompanyName
	}

	hidden [bool]IsDefaultApp(){
		return $this.AppName -eq $this.Cfg.DefaultAppName
	}

	hidden [void]ResolveCompanyPath([string]$CompanyName){
		# Split on backslash to detect true nested hierarchy
		$parts = $CompanyName -split '\\'
		# Start at the root (e.g., HKCU\Software)
		if($null -eq $this.HiveSoftwareRoot){
			$this.HiveSoftwareRoot = [W32Reg]::$($this.Hive).OpenSubKey($this.cfg.RegistryRoot,$true)
		}
		$key = $this.HiveSoftwareRoot
		foreach ($part in $parts){
			if ([string]::IsNullOrWhiteSpace($part)){
				throw ('Invalid company name segment in ({0}).' -f $CompanyName)
			}
			# Create or open each level
			$key = $key.CreateSubKey($part)
			if ($null -eq $key){
				throw ('Failed to create or open registry key segment ({0}).' -f $part)
			}
		}
	}
#endregion

#region Methods
	[void]SetRegistryForWrite([bool]$Mode){$this.WriteMode  = $Mode}

	[bool]IsWriteable(){return $this.WriteMode }

	[string]GetHiveName(){return $this.Hive}

	[string]GetInitialCompany(){return $this.Cfg.CompanyName}

	[void]SetInitialCompany([string]$Name){
		$ErrMsg = @(
			@('SetInitialCompany Request Ignored, Initial Company Name Assigned & Locked!','Yellow'),
			@('You may set the active company name using the CompanyName Property.','White')
		)
		if([string]::IsNullOrWhiteSpace($Name)){
			throw [ArgumentException]::new(
				[AppRegistry]::MsgStore([MsgID]::MustBeNonEmptyString) -f 'CompanyName'
			)
		}
		if(!$this.Cfg.CompanyLock){
			$this.Cfg.CompanyName = $Name 
			$this.Cfg.CompanyLock = $true
			$this.CompanyName = $Name
		}else{For($C = 0;$C -le 1; $C++){Write-Host $ErrMsg[$C][0] -ForegroundColor $ErrMsg[$C][1]}}
	}

	[void]RestoreInitialCompany(){$this.CompanyName = $this.Cfg.CompanyName}

	[void]SetBaseKey([string]$Company,[string]$AppName){
		$this.CompanyName = $Company
		$this.AppName     = $AppName
		$this.ValidateBaseKeyNames()
	}

	[W32RegKey]OpenCompanyKey(){
		$this.ValidateBaseKeyNames()
		$Root = $this.GetCompanyRootPath()
		$this.ResolveCompanyPath($this.CompanyName)
		return [W32Reg]::$($this.Hive).OpenSubKey($Root,$this.WriteMode )
	}

	[W32RegKey]OpenApplicationKey(){
		$this.ValidateBaseKeyNames
		$this.ResolveCompanyPath($this.CompanyName)
		$Root = $this.GetAppRootPath()
		return [W32Reg]::$($this.Hive).OpenSubKey($Root, $this.WriteMode )
	}

	[void]SetValue([string]$Name,$Value,[W32RegKind]$Kind){
		if(-not $this.WriteMode ){
			throw [InvalidOperationException]::new('This registry context is read-only.')
		}

		if ($Value -is [object[]]) {
		    if($Value.Count -eq 0){Return} #Empty Array			
			else{
				[string[]]$Str = $()
				foreach($Item in $Value){
					$Str += $Item.ToString() #Coerce to string
				}
				$Value = $Str
			}		
		}

		switch ($Kind) {
			([W32RegKind]::String) {
				if ($Value -isnot [string]) {
					throw [ArgumentException]::new(
						[AppRegistry]::MsgStore([MsgID]::MustBeRegistryValueKind) -f 'string','String'
					)
				}
				break
			}
			([W32RegKind]::DWord) {
				if ($Value -isnot [int]) {
					throw [ArgumentException]::new(
						[AppRegistry]::MsgStore([MsgID]::MustBeRegistryValueKind) -f 'Int32','DWord'
					)
				}
				break
			}
			([W32RegKind]::QWord) {
				if ($Value -isnot [long]) {
					throw [ArgumentException]::new(
						[AppRegistry]::MsgStore([MsgID]::MustBeRegistryValueKind) -f 'Int64','QWord'
					)
				}
				break
			}
			([W32RegKind]::Binary) {
				# Dutchman patch: expands in two directions
				if ($Value -is [bool]) {
					[Byte[]]$Value = if ($Value) {[byte]1} else {[byte]0}
				}
				elseif ($Value -isnot [byte[]]) {
					throw [ArgumentException]::new('Binary values must be Boolean or Byte[].')
				}
				break
			}
			([W32RegKind]::MultiString) {
				if ($Value -isnot [string[]]) {
					throw [ArgumentException]::new(
						[AppRegistry]::MsgStore([MsgID]::MustBeRegistryValueKind) -f 'string[]','MultiString'
					)
				}
				break
			}
		}
		$key = $this.OpenApplicationKey()
		if ($null -eq $key){
			$this.CreateApplicationKey()
			$key = $this.OpenApplicationKey()
		}
		$key.SetValue($Name, $Value, $Kind)
	}

	[bool]ParseBinaryAsBoolean($Value){
		if($Value -isnot [byte[]] -or $Value.Length -ne 1){
			throw [ArgumentException]::new('Value is not a single-byte REG_BINARY boolean.')
		}
		return $Value[0] -eq 1
	}

	[void]CreateCompanyKey(){
		$this.ValidateBaseKeyNames()
		$root = $this.GetCompanyRootPath()
		$this.ResolveCompanyPath($this.CompanyName)
		[void][W32Reg]::$($this.Hive).CreateSubKey($root)
	}

	[void]CreateApplicationKey(){
		$this.ValidateBaseKeyNames()
		$root = $this.GetAppRootPath()
		$this.ResolveCompanyPath($this.CompanyName)
		[void][W32Reg]::$($this.Hive).CreateSubKey($root)
	}

	[bool]CompanyKeyExists(){
		$RK = [W32Reg]::$($this.Hive).OpenSubKey($this.GetCompanyRootPath())
		return $null -ne $RK
	}

	[bool]ApplicationKeyExists(){
		$RK = [W32Reg]::$($this.Hive).OpenSubKey($this.GetAppRootPath())
		return $null -ne $RK
	}

	[bool]ValueExists([string]$Name){
		$RK = [W32Reg]::$($this.Hive).OpenSubKey($this.GetAppRootPath())
		if ($null -eq $RK) {return $false}
		$RV = $RK.GetValue($Name)
		return $null -ne $RV
	}

	[void]FlushApplicationKey(){
		$CompanyRoot = $this.GetCompanyRootPath()
		$RK = [W32Reg]::$($this.Hive).OpenSubKey($CompanyRoot, $true)
		if($null -eq $RK){
			throw [AppRegistry]::MsgStore([MsgID]::ApplicationKeyNonExistant) -f `
				[AppRegistry]::GetHiveAbbrev($this.Hive),$CompanyRoot
		}
		$RK.DeleteSubKey($this.AppName)
		if($RK.SubKeyCount -eq 0){
			$Root = $this.Cfg.RegistryRoot
			$Top  = [W32Reg]::$($this.Hive).OpenSubKey($Root, $true)
			$Top.DeleteSubKey($this.CompanyName)
		}
	}

	[void]RemoveValue([string]$Name){
		$this.ValidateBaseKeyNames()
		$Root = $this.GetAppRootPath()
		$RK = [W32Reg]::$($this.Hive).OpenSubKey($Root,$true)
		if($null -eq $RK){
			throw [AppRegistry]::MsgStore([MsgID]::ApplicationKeyNonExistant) -f `
				[AppRegistry]::GetHiveAbbrev($this.Hive),$Root
		}

		$RV = $RK.GetValue($Name)
		if($null -eq $RV){
			throw 'Application Value {0}:{1}\{2} is Non-Existant' -f `
				[AppRegistry]::GetHiveAbbrev($this.Hive),$Root,$Name
		}

		$RK.DeleteValue($Name,$true)
	}

	[void]ClearBaseKeys(){
		$this.ApplicationKey = $null
		$this.CompanyKey     = $null
	
	}

	[void]LoadBaseKeys(){
		$this.ValidateBaseKeyNames()
		$this.ResolveCompanyPath($this.CompanyName)
		$this.ApplicationKey =
			[W32Reg]::$($this.Hive).OpenSubKey($this.GetAppRootPath(), $this.WriteMode )
		$this.CompanyKey =
			[W32Reg]::$($this.Hive).OpenSubKey($this.GetCompanyRootPath(), $this.WriteMode )
	}

	[void]RefreshBaseKeys(){$this.LoadBaseKeys()}

	[void]ResetBaseKeys(){$this.LoadBaseKeys()}

	[void]ValidateBaseKeyNames(){
		if([string]::IsNullOrWhiteSpace($this.CompanyName) -or $this.IsDefaultCompany()){
			throw [ArgumentException]::new(
				[AppRegistry]::MsgStore([MsgID]::MustBeNonEmptyString) -f 'CompanyName'
			)
		}
		if([string]::IsNullOrWhiteSpace($this.AppName) -or $this.IsDefaultApp()){
			throw [ArgumentException]::new(
				[AppRegistry]::MsgStore([MsgID]::MustBeNonEmptyString) -f 'AppName'
			)
		}
	}
#endregion
}

Class BinaryPair{
	[string]$Pos #Positive
	[string]$Neg #Negative

	BinaryPair(){}
	BinaryPair([string]$Pos,[string]$Neg){
		$this.Pos  = $Pos
		$this.Neg  = $Neg
	}

	static [BinaryPair]FromStrings([string]$Pos,[string]$Neg){
		return [BinaryPair]::new($Pos,$Neg)
	}
}

Class ResolveBinaryPairs{
	#region Fields
	hidden [byte]$BpTrue
	hidden [byte]$BpFalse
	hidden [BinaryPair[]]$DefaultBinaryPairs
	hidden [BinaryPair[]]$UserDialect
	hidden [System.Collections.Generic.SortedDictionary[string,byte]]$BinaryDictionary
	#endregion

	#region Constructors
	ResolveBinaryPairs(){
		# Canonical byte values
		$this.BpTrue  = 1
		$this.BpFalse = 0

		# Base dialect (never mutated at runtime)
		$this.DefaultBinaryPairs = @(
			[BinaryPair]::new('True','False'),
			[BinaryPair]::new('On','Off'),
			[BinaryPair]::new('Yes','No'),
			[BinaryPair]::new('Lock','Unlock'),
			[BinaryPair]::new('1','0')
		)

		# Lookup dictionary (ordered)
		$this.BinaryDictionary =
			[System.Collections.Generic.SortedDictionary[string,byte]]::new(
				[System.StringComparer]::InvariantCultureIgnoreCase
		)

		$this.SeedDictionary()
		$this.UserDialect = [BinaryPair[]]@()
	}
	#endregion

	#region Public API
	hidden [void]SeedDictionary(){
		$this.BinaryDictionary.Clear()
		foreach ($pair in $this.DefaultBinaryPairs) {
			$this.AddPair($pair)
		}
	}

	hidden [void]SeedDialect(){
		$this.BinaryDictionary.Clear()
		foreach ($pair in $this.UserDialect) {
			$this.AddPair($pair)
		}
	}

	[void]InstallDialect([BinaryPair[]]$Dialect){
		$this.UserDialect = $Dialect
		$this.SeedDialect()
	}

	[void]AddPair([BinaryPair]$Pair){
		if([string]::IsNullOrWhiteSpace($Pair.Pos) -or [string]::IsNullOrWhiteSpace($Pair.Neg)){
			throw 'Keywords cannot be null or empty.'
		}

		if ($this.BinaryDictionary.ContainsKey($Pair.Pos) -or
			$this.BinaryDictionary.ContainsKey($Pair.Neg)
		) { Write-Host ('AddPair: Duplicate Key(s) Detected, {0} Ignored' -f $Pair) -ForegroundColor Yellow}
		else {
			$this.BinaryDictionary.Add($Pair.Pos, $this.BpTrue)
			$this.BinaryDictionary.Add($Pair.Neg, $this.BpFalse)
		}
	}

	[void]AddPairs([BinaryPair[]]$Pairs){
		foreach ($Pair in $Pairs){$this.AddPair($Pair)}
	}

	[byte]Resolve([string]$InputValue){
		[byte]$out = $null
		if([string]::IsNullOrWhiteSpace($InputValue)){
			throw 'Input cannot be null or empty.'
		}
		if ($this.BinaryDictionary.TryGetValue($InputValue, [ref]$out)){
			return $out
		}
		throw ('Unrecognized binary keyword: ({0}).' -f $InputValue)
	}

	[bool]TryResolve([string]$InputValue,[ref]$Result){
		[byte]$out = $null
		if([string]::IsNullOrWhiteSpace($InputValue)){
			$Result.Value = 0
			return $false
		}
		if($this.BinaryDictionary.TryGetValue($InputValue,[ref]$out)){
			$Result.Value = $out
			return $true
		}
		$Result.Value = 0
		return $false
	}

	[void]RemovePair([BinaryPair]$pair){
		$this.BinaryDictionary.Remove($pair.Pos)
		$this.BinaryDictionary.Remove($pair.Neg)
	}

	[void]RemovePairs([BinaryPair[]]$pairs){
		foreach ($pair in $pairs){$this.RemovePair($pair)}
	}

	[void]ResetToDefault(){$this.SeedDictionary()}

	[void]RestoreDialect(){$this.SeedDialect()}
	#endregion

	#region Query Methods
	[Byte]GetCanonicalTrueByte(){return $this.BpTrue}
	[Byte]GetCanonicalFalseByte(){return $this.BpFalse}
	[BinaryPair[]]GetDefaultBinaryPairs(){return $this.DefaultBinaryPairs}
	[BinaryPair[]]GetDialect(){return $this.UserDialect}
	[System.Collections.Generic.SortedDictionary[string,byte]]CloneLookupDictionary(){return $this.BinaryDictionary}
	#endregion
}

#region Sample Calls
<# 
Import-Module .\AppRegistryClass.ps1 -force
$ARN = [AppRegistry]::new('Blueflame Dynamics','CurrentUser',$true)
$ARN.AppName = 'PS Audio Player'
$ARN.LoadBaseKeys()
$ARN.ApplicationKey.GetValueNames()
$ARN.ApplicationKey.GetValue('Playlist')

# ResolveBinaryPair
Import-Module .\AppRegistryClass.ps1 -force
$BP = [ResolveBinaryPairs]::new()
$BP.AddPair([BinaryPair]::new('+','-'))
$BP.AddPair([BinaryPair]::new('Jahvol','Nien'))
$BP.AddPair([BinaryPair]::new('Yin','Lin'))
$BP.AddPair([BinaryPair]::new('Yay','Nay'))
$BP.CloneLookupDictionary()
$BP.RemovePair([BinaryPair]::new('Yin','Lin'))
$BP.CloneLookupDictionary()
[byte]$result = 8
$success = $BP.TryResolve('Nien', [ref]$result)
'Success: {0}{2}Result:  {1}{2}Success: 1/0 = True/False Key' -f $success,$result,"`r`n"
#>
#endregion