<#
.NOTES
-------------------------------------
Name:    SplashScreen.ps1
Version: 4.0 - 2026/02/08
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
Update History:
	V4.0 - 2026/02/08 - Merged SplashScreen & SplashScreenW32 codebases.
	V1.0 - 2025/11/20 - Major Update adds progress updates & Branches off as SplashScreenW32.
	V3.0 - 2022/10/22 - Rewrite, Introduced SplashRunspace Class
	V1.0 - 2018/06/08 - Original Static Splash Screen

.SYNOPSIS
This script launches a Winform Splash Screen on a Background Thread.
There are 3 functions, Show-SplashScreen, Close-SplashScreen & Update-Splash. This version uses
a Class to both define & create a SplashRunspace object used by the functions.
As a cmdlet you may run a demo of a Splash Screen but when imported as a
module you may call the all functions directly.
----------------------------------------------------------------------------------------

.DESCRIPTION
There are 3 primary functions, Show-SplashScreen, Close-SplashScreen & Update-Splash, When calling 
Show-SplashScreen 5 parameters allow specifing the Calling Application Name, An Image 
to Display, and Optionally the form background color, the PictureBox background color
& a switch which causes the image to be used as a background image. By default the 
image is displayed inside a PictureBox Control.
.LINK
Show-SplashScreen
.LINK
Close-SplashScreen
.LINK
Update-Splash
---------------------------------------------------------------------------------------- 
.Parameter Demo
An optional switch that will cause a demo splash screen to be displayed.
.Parameter Delay
An optional value in seconds that may be used to alter the default 30 second demo duration.
.Parameter Mode
Optional, Specifies mode of operation
.Parameter ImageSize
Optional sets the square image size for the demo.
.Parameter PbSizeMode
Optional sets the [Windows.Forms.PictureBoxSizeMode] for the demo.
.Parameter BgImage 
Optional switch causes the image to be background image
.Parameter BGILayout
Optional BackgroundImageLayout.
.Parameter ShowPbText
Optional, Show Progress percentage in Modern mode
#>
[CmdletBinding()]
param(
	[Parameter()][Switch]$Demo,
	[Parameter()][Int]$Delay=30,
	[Parameter()][ValidateSet('Legacy','Modern')][String]$Mode='Legacy',
	[Parameter()][Int]$ImageSize=256,
	[Parameter()][System.Windows.Forms.PictureBoxSizeMode]$PbSizeMode=[Windows.Forms.PictureBoxSizeMode]::Zoom,
	[Parameter()][System.Windows.Forms.ImageLayout]$BGILayout=[Windows.Forms.ImageLayout]::Zoom,
	[Parameter()][Switch]$BgImage,
	[Parameter()][Switch]$ShowPbText)

<#
Declare the Custom Runspace Object used by SplashScreen methods as
a Script level variable and a value of $Null. This variable will be loaded with a new
SplashRunspace object each time you call Show-SplashScreen and this script level 
variable is null otherwise Show-SplashScreen exits assuming a splash screen is 
already running. As part of Close-SplashScreen this variable wiil be reset to $Null
this prevents multiple calls to Show-SplashScreen without a call to Close-SplashScreen.
#>
$Script:SplashRunspace = $Null

#region Class SplashRunspace
<#
.DESCRIPTION
There are 3 top-level properties: HashTable is a Synchronized hash table, thread safe for inter-thread communications.
The HashTable.Flag is a boolean when true, a splash screen may be shown, setting this value to false will cause the 
splash screen to close and the thread to be released. HashTable.SplashIsLoaded is an integer value when value 
is zero the form initialization occurs, once initialization has occured this is incremented to 1 to prevent the form
being redefined in subsequent passes of a do-while loop. Once incremented the splash screen is displayed via the
ShowDialog() method. The splash screen remains active until the Close-SplashScreen function signals it to close
by setting HashTable.Flag to false. The second top-level property is Powershell a [Powershell] type object
representing the separate thread where the splash screen script will run. This property inherits all the attributes
of the [Powershell] object it contains. The last top-level property is Handle, this is the [IAsyncResult] object
returned when the background thread is fired by a call to the Splashrunspace.Powershell.BeginInvoke() method. This
objects properties provide status information about the background thread and is passed back when Close-SplashScreen 
is invoked via a SplashRunspace.PowerShell.EndInvoke() after the splash screen has closed and the Handle.IsCompeleted 
property is true. The SplashRunspace.PowerShell property is disposed and the $Script:SplashRunspace variable is
reset to $Null. When you create a New SplashRunspace object you must pass the Name of the SplashRunspace.HashTable
in the associated splash screen script.
#>
Class SplashRunspace {
	#Properties
	[HashTable]$HashTable
	[PowerShell]$Powershell
	[IAsyncResult]$Handle

	#New Method Constructor
	SplashRunspace([String]$CommunicationsHashTableName){
		$This.HashTable = [HashTable]::Synchronized(@{})
		$This.HashTable.Flag = $True
		$This.HashTable.SplashIsLoaded = 0
		$This.HashTable.ProgressValue = 0
		$This.HashTable.ProgressMessage = "Starting..."
		$This.Powershell = [PowerShell]::Create()
		$This.Powershell.Runspace = [RunspaceFactory]::CreateRunspace()
		$This.Powershell.Runspace.Open()
		$This.Powershell.Runspace.SessionStateProxy.SetVariable($CommunicationsHashTableName,$This.HashTable)
	}
}
#endregion

#region Functions
<#
.NOTES
-------------------------------------
Name:    Invoke-Ternary
Version: 1.0 - 2025/11/20
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
Update History:
	V1.0 - 2025/11/20 - Original Powershell 5.1+ Ternary operator.
.SYNOPSIS
	Powershell 5.1+ Ternary operator with ?: alias.
	Syntax: $Mode = ?: (<Condition>) <True> <False> 
#>
Function Invoke-Ternary {
	param(
		[Parameter(Mandatory, Position=0)]$Condition,
		[Parameter(Mandatory, Position=1)]$IfTrue,
		[Parameter(Mandatory, Position=2)]$IfFalse)

	# If condition is a scriptblock, evaluate it
	$condResult = if ($Condition -is [scriptblock]) { & $Condition } else { $Condition }

	if ($condResult) {
		if ($IfTrue -is [scriptblock]) { & $IfTrue } else { $IfTrue }
	} else {
		if ($IfFalse -is [scriptblock]) { & $IfFalse } else { $IfFalse }
	}
}
Set-Alias -Name ?: -Value Invoke-Ternary

[Int]$TotalWeight = 0
<#
.NOTES
	Name:	Function Update-Splash
	Version:	1.0 - 2025/11/29
	Author:		Randy Turner
	Email:		turner.randy21@yahoo.com
	Created:	2025/11/29
.SYNOPSIS
	Function to update splash Message & ProgressBar.Value.
	The ProgressBar.Value Total may be reset to zero by
	calling this method without parameters.
.PARAMETER Message Alias: M
	Progress Message Text.
.PARAMETER Weight Alias: W
	ProgressBar.Value Incremental Value.
#>
Function Update-Splash{
[CmdletBinding()]
	param(
		[Parameter()][String]$Message=$null,
		[Parameter()][Int]$Weight=0,
		[Parameter()][Switch]$ShowWeights)

		if($Null -eq $Message -and $Weight -eq 0){
			#Reset Requested!
			$Script:TotalWeight = 0
		}Else{
		$Script:TotalWeight += $Weight
		$SplashRunspace.HashTable.ProgressMessage = $Message
		$SplashRunspace.HashTable.ProgressValue   = $Script:TotalWeight
		}

		if($ShowWeights){'TotalWeight: {0} Weight: {1}' -f $Script:TotalWeight, $Weight}
		if($Script:TotalWeight -eq 100){$Script:TotalWeight = 0}
		Start-Sleep -Seconds 5
}

<#
.NOTES
-------------------------------------
Name:    Function Show-SplashScreen
Version: 2.5 - 2026/02/08
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
.SYNOPSIS
This function Displays a Splash Screen on a Background Thread.
V2.5 includes a rewrite of SplashScript to support 2 modes and flatten
logic to reduce cognative load, easeing maintainance.
----------------------------------------------------------------------------------------
.DESCRIPTION
This function has 1 parameter a custom PSObject with An Image to Display.
Optionally, Calling Application Name, Form background color, PictureBox background color,
BackgroundImageLayout to Apply & a boolean which causes the image to be used as a 
background image. By default the image is displayed inside a PictureBox Control.
When the ProgressBar is started in Marquee style, it switches to Continuous style 
once progress values appear like a Windows OS ProgressBar.
This Script is compatible with Powershell 5.1+.
This function communicates with the background thread via a synchronized hash table
stored in the SplashRunspace object. This allows Update‑Splash and Close‑SplashScreen
to interact with the UI safely.”
---------------------------------------------------------------------------------------- 
.PARAMETER Config
Required, This is a Custom PSObject used to pass parameter values.
.EXAMPLE
$cfgObj = [PsObject]@{
	AppName    = 'SplashScreen Demo'
	Image = $SSIcon 
	FormBackColor = [Drawing.Color]::CornflowerBlue
	Message = 'Initializing modules...'
	ProgressBarStyle = [Windows.Forms.ProgressBarStyle]::Marquee
}
$CfgObj|Show-SplashScreen 
Displays the Splash Screen using a PictureBox
#>
Function Show-SplashScreen {
param([Parameter(ValueFromPipeline, Mandatory)][PSObject]$Config)

begin {
		if ($null -ne $Script:SplashRunspace) { return }
	}
process {
	$Image         = $Config.Image # Required Parameter
	$Mode          = ?: ($Null -ne $Config.Mode) {$Config.Mode} 'Legacy'
	$AppName       = ?: ($Null -ne $Config.AppName) {$Config.AppName} 'Application'
	$FormBackColor = ?: ($Config.FormBackColor) $Config.FormBackColor [Drawing.SystemColors]::Control
	$PicBackColor  = ?: ($Null -ne $Config.PicBackColor) {$Config.PicBackColor} {[Drawing.SystemColors]::Control}
	$PbSizeMode    = ?: ($Null -ne $Config.PbSizeMode) {$Config.PbSizeMode} {[Windows.Forms.PictureBoxSizeMode]::Zoom}
	$BGILayout     = ?: ($Null -ne $Config.BGILayout) {$Config.BGILayout} {[Windows.Forms.ImageLayout]::Zoom}
	$Message       = ?: ($Null -ne $Config.Message) {$Config.Message} {("{0} Loading, Please Wait ..." -f $AppName)}
	$ProgressBarForeColor = ?: ($Null -ne $Config.ProgressBarForeColor) {$Config.ProgressBarForeColor} {[Drawing.SystemColors]::Highlight}
	$ProgressBarBackColor = ?: ($Null -ne $Config.ProgressBarBackColor) {$Config.ProgressBarBackColor} {[Drawing.SystemColors]::Control}
	$ProgressBarStyle = if ($Null -ne $Config.ProgressBarStyle) { $Config.ProgressBarStyle } else { [Windows.Forms.ProgressBarStyle]::Continuous }
	$BgImage = $Config.BgImage
	$ShowPbText = $Config.ShowPbText

	if (-not $Image) { throw "Image is required in Config object." }

$SplashScript = {
	param(
		[Parameter(Mandatory)][String]$Mode,
		[Parameter(Mandatory)][Drawing.Image]$Image,
		[Parameter()][String]$AppName = 'Application',
		[Parameter()][Drawing.Color]$FormBackColor = [Drawing.SystemColors]::Control,
		[Parameter()][Drawing.Color]$PicBackColor = [Drawing.SystemColors]::Control,
		[Parameter()][Windows.Forms.PictureBoxSizeMode]$PbSizeMode = [Windows.Forms.PictureBoxSizeMode]::Zoom,
		[Parameter()][Windows.Forms.ImageLayout]$BGILayout = [Windows.Forms.ImageLayout]::Zoom,
		[Parameter()][Drawing.Color]$ProgressBarForeColor = [Drawing.SystemColors]::Highlight,
		[Parameter()][Drawing.Color]$ProgressBarBackColor = [Drawing.SystemColors]::Control,
		[Parameter()][Windows.Forms.ProgressBarStyle]$ProgressBarStyle = [Windows.Forms.ProgressBarStyle]::Continuous,
		[Parameter()][Switch]$BgImage,
		[Parameter()][Switch]$ShowPbText)

	Function New-Splash{       
		# --- Form + Controls ----------------------------------------------------
		$UI = [PSCustomObject]@{
			Splash       = [Windows.Forms.Form]::New()
			ProgressBar  = [Windows.Forms.ProgressBar]::New()
			LblMsg       = [Windows.Forms.Label]::New()
			PctMsg       = [Windows.Forms.Label]::New()
			PSCoreFont   = [Drawing.Font]::New('Segoe UI', 9, [Drawing.FontStyle]::Regular)
			Timer        = [Windows.Forms.Timer]::New()
			SharedWidth  = 0
			PicBox       = $null}
		# --- Splash --------------------------------------------------------
		$UI.Splash.Size            = [Drawing.Size]::New(475, 325)
		$UI.Splash.StartPosition   = [Windows.Forms.FormStartPosition]::CenterScreen
		$UI.Splash.Font            = [Drawing.Font]::New('Arial', 12, [Drawing.FontStyle]::Regular)
		$UI.Splash.BackColor       = $FormBackColor
		$UI.Splash.FormBorderStyle = [Windows.Forms.BorderStyle]::None
		$UI.Splash.ShowInTaskbar   = $False
		$UI.SharedWidth = $UI.Splash.Width - 25
		# --- ProgressBar --------------------------------------------------------
		$UI.ProgressBar.Style     = $ProgressBarStyle
		$UI.ProgressBar.ForeColor = $ProgressBarForeColor
		$UI.ProgressBar.BackColor = $ProgressBarBackColor
		$UI.ProgressBar.Parent    = $UI.Splash
		$UI.ProgressBar.Minimum   = 0
		$UI.ProgressBar.Maximum   = 100
		$UI.ProgressBar.Visible   = $True
		$UI.ProgressBar.Size      = [Drawing.Size]::New($UI.SharedWidth, 20)
		$UI.ProgressBar.Location  = [Drawing.Point]::New(13, $UI.Splash.Height - ($UI.ProgressBar.Height + 10))
		# --- Progress Label --------------------------------------------------------------
		if($ShowPbText){
		    $UI.PctMsg.Parent      = $UI.Splash
		    $UI.PctMsg.Size        = [Drawing.Size]::New(35, 20)
		    $UI.PctMsg.Location    = $UI.ProgressBar.Location
			$UI.ProgressBar.Width -= $UI.PctMsg.Width
			$UI.ProgressBar.Left  += $UI.PctMsg.Width
		    $UI.PctMsg.BorderStyle = [Windows.Forms.BorderStyle]::None
		    $UI.PctMsg.BackColor   = [Drawing.Color]::Transparent
		    $UI.PctMsg.ForeColor   = [Drawing.Color]::Black
		    $UI.PctMsg.Font        = $UI.PSCoreFont
		    $UI.PctMsg.Text        = "{0}%" -f $UI.ProgressBar.Value
		    $UI.PctMsg.TextAlign   = [Drawing.ContentAlignment]::MiddleCenter 
		    $UI.PctMsg.Visible     = $True}
		# --- Msg Label --------------------------------------------------------------
		$UI.LblMsg.Parent      = $UI.Splash
		$UI.LblMsg.Size        = [Drawing.Size]::New($UI.SharedWidth, 20)
		$UI.LblMsg.Location    = [Drawing.Point]::New(13, $UI.ProgressBar.Top - ($UI.LblMsg.Height + 5))
		$UI.LblMsg.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
		$UI.LblMsg.BackColor   = [Drawing.SystemColors]::Control
		$UI.LblMsg.Font        = $UI.PSCoreFont
		$UI.LblMsg.Text        = "{0} Loading, Please Wait ..." -f $AppName
		$UI.LblMsg.Visible     = $True
		# --- Image / PictureBox -------------------------------------------------
		if ($BgImage) {
			$UI.Splash.BackgroundImage       = $Image
			$UI.Splash.BackgroundImageLayout = $BGILayout
		}
		else {
			$UI.PicBox = [Windows.Forms.PictureBox]::New()
			$UI.PicBox.Parent      = $UI.Splash
			$UI.PicBox.Size        = [Drawing.Size]::New($UI.SharedWidth, $UI.LblMsg.Top - 18)
			$UI.PicBox.Location    = [Drawing.Point]::New(13, 13)
			$UI.PicBox.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
			$UI.PicBox.Image       = $Image
			$UI.PicBox.Visible     = $True
			$UI.PicBox.BackColor   = $PicBackColor
			$UI.PicBox.SizeMode    = $PbSizeMode
		}
		return $UI 
	}
	Function Invoke-Timer{
		# --- Timer --------------------------------------------------------------
		$UI.Timer.Interval = 500
		$Timer_Tick = {
			if ($SyncHash.Flag -eq $false) {
				$UI.Splash.Close()
				return
			}

			if ($Mode -eq 'Legacy') {
				# --- Legacy fixed progress bar behavior ---
				if ($UI.ProgressBar.Value -ge $UI.ProgressBar.Maximum) {
					$UI.ProgressBar.Value = $UI.ProgressBar.Minimum
				}
				$UI.ProgressBar.Value++
				# --- Legacy fixed message ---
				$UI.LblMsg.Text = "{0} Loading, Please Wait ..." -f $AppName
				# --- Legacy fixed style ---
				$UI.ProgressBar.Style = [Windows.Forms.ProgressBarStyle]::Continuous
			}
			else {
				# --- Modern dynamic progress behavior ---
				if ($SyncHash.ProgressValue -ge $UI.ProgressBar.Minimum -and
					$SyncHash.ProgressValue -le $UI.ProgressBar.Maximum) {
					$UI.ProgressBar.Value = $SyncHash.ProgressValue
					$UI.PctMsg.Text = "{0}%" -f $UI.ProgressBar.Value
					if ($UI.ProgressBar.Style -eq [Windows.Forms.ProgressBarStyle]::Marquee -and
						$SyncHash.ProgressValue -lt $UI.ProgressBar.Maximum) {
						$UI.ProgressBar.Style = [Windows.Forms.ProgressBarStyle]::Continuous
					}
				}
				else {
					$UI.ProgressBar.Style = [Windows.Forms.ProgressBarStyle]::Marquee
				}
				# --- Modern dynamic message ---
				if ($UI.ProgressBar.Value -gt 0) {
					$UI.LblMsg.Text = $SyncHash.ProgressMessage
				}
			}
			[Windows.Forms.Application]::DoEvents()
		}
		$UI.Timer.Add_Tick($Timer_Tick)
		$UI.Timer.Enabled = $True
		$UI.Timer.Start()			
	}
	Function Enable-SplashDragging{
		# --- Win32 Dragging -----------------------------------------------------
		$Splash_Shown = {
			$UI.Splash.Activate()
			$UI.Splash.TopMost = $True
			$UI.Splash.BringToFront()
			$UI.Splash.TopMost = $False
		}
		$UI.Splash.Add_Shown($Splash_Shown)

		$Splash_MouseDown = {
			if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
				[Win32]::ReleaseCapture() | Out-Null
				[Win32]::SendMessage($UI.Splash.Handle, $WM_NCLBUTTONDOWN, $HTCAPTION, 0)
			}
		}
		$UI.Splash.Add_MouseDown($Splash_MouseDown)
		foreach ($ctrl in $UI.Splash.Controls) {
			$ctrl.Add_MouseDown($Splash_MouseDown)
		}	
	}

	while ($SyncHash.Flag) {
		if ($SyncHash.SplashIsLoaded -eq 0) {
			Add-Type -AssemblyName System.Windows.Forms
#region Win32 Dragging
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32 {
	[DllImport("user32.dll")]
	public static extern bool ReleaseCapture();

	[DllImport("user32.dll")]
	public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
}
"@
$WM_NCLBUTTONDOWN = 0xA1
$HTCAPTION        = 0x2
#endregion
			$UI = New-Splash
			Invoke-Timer
		    Enable-SplashDragging
			# --- Finalize -----------------------------------------------------------
			$InitialFormWindowState = $UI.Splash.WindowState
			$UI.Splash.Add_Load({ $This.WindowState = $InitialFormWindowState })
			$SyncHash.SplashIsLoaded++
			$UI.Splash.ShowDialog()
		}
	}
}

	# --- Set SplashRunspace Parameters & Invoke ----------------------
	$Script:SplashRunspace = [SplashRunspace]::New('SyncHash')
	[void]$SplashRunspace.Powershell.AddScript($SplashScript)
	[void]$SplashRunspace.Powershell.AddParameter('Mode',$Mode)
	[void]$SplashRunspace.Powershell.AddParameter('AppName',$AppName)
	[void]$SplashRunspace.Powershell.AddParameter('Image',$Image)
	[void]$SplashRunspace.Powershell.AddParameter('FormBackColor',$FormBackColor)
	[void]$SplashRunspace.Powershell.AddParameter('PicBackColor',$PicBackColor)
	[void]$SplashRunspace.Powershell.AddParameter('PbSizeMode',$PbSizeMode)
	[void]$SplashRunspace.Powershell.AddParameter('ProgressBarForeColor',$ProgressBarForeColor)
	[void]$SplashRunspace.Powershell.AddParameter('ProgressBarBackColor',$ProgressBarBackColor)
	[void]$SplashRunspace.Powershell.AddParameter('ProgressBarStyle',$ProgressBarStyle)
	if($ShowPbText){[void]$SplashRunspace.Powershell.AddParameter('ShowPbText')}
	if($BgImage){
	    [void]$SplashRunspace.Powershell.AddParameter('BgImage')
	    [void]$SplashRunspace.Powershell.AddParameter('BgiLayout',$BGILayout)
	}
	$SplashRunspace.Handle = $SplashRunspace.PowerShell.BeginInvoke()
	} #end Process
}

<#
.NOTES
-------------------------------------
Name:    Function Close-SplashScreen
Version: 1.0 - 03/25/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This function Closes a Splash Screen on a Background Thread & Disposes of the Thread.
----------------------------------------------------------------------------------------
.DESCRIPTION
This function Closes a Splash Screen 
---------------------------------------------------------------------------------------- 
.EXAMPLE
Close-SplashScreen
Closes the Splash Screen
#>
Function Close-SplashScreen{
$SplashRunspace.HashTable.Flag = $False    #Close Screen
if($SplashRunspace.Handle.IsCompleted)
	{#Dispose of Thread
	$SplashRunspace.PowerShell.EndInvoke($SplashRunspace.Handle)
	$SplashRunspace.PowerShell.Dispose()
	$Script:SplashRunspace = $Null
	}
}
#endregion

if($Demo.IsPresent){
Function Get-Icon{
	# Default Image (Powershell Icon) used for Demo/Testing, Index by OS
	Import-Module -Name .\Exists.ps1 -Force
	Import-Module -Name .\OSInfoLib.ps1 -Force
	$IcoIdx = If((Get-OSVersion).IsWin11){312}Else{311}
	$DLLPath = '.\BlueflameDynamics.IconTools.dll'
	if((Test-Exists -Mode File -Location $DLLPath) -eq $True){
		Add-Type -Path $DLLPath
		Start-Sleep -Seconds 1
		$Icon = [BlueflameDynamics.IconTools]::ExtractIcon(
			$Env:WinDir+'\System32\ImageRes.dll',$IcoIdx,$ImageSize)
	}
	$Icon
}

Function Start-Demo {
	[CmdletBinding()]
	param(
		[Parameter()][Int]$Delay=30,
		[Parameter()][Array]$Steps = @(
			@{Name = 'Loading configuration'; Weight = 20},
			@{Name = 'Initializing modules';  Weight = 30},
			@{Name = 'Connecting services';   Weight = 30},
			@{Name = 'Finalizing startup';    Weight = 20}
		)
	)

	# Confugure splash
	$config = New-Object PSObject -Property @{
	AppName = ?: ($Mode -eq 'Modern') 'SplashScreen Demo' {$Null}
	Mode = $Mode
	Image = $SSIcon
	ShowPbText = $ShowPbText
	FormBackColor = [Drawing.Color]::CornflowerBlue
	BgImage = $BgImage
	BgiLayout = $BGILayout
	ImageSize = $ImageSize
	#Message = 'Initializing modules...'
	# Starts in marquee, switches to continuous once values appear,
	ProgressBarStyle = [Windows.Forms.ProgressBarStyle]::Marquee 
	#Uncomment to Override Powershell Defaults
	#ProgressBarForeColor = [Drawing.Color]::LimeGreen
	#ProgressBarBackColor = [Drawing.Color]::ForestGreen
	<#
	[Windows.Forms.ProgressBarStyle]::Blocks is ignored by some Powershell Hosts like the ISE.
	This is because 'Blocks' is the default ProgressBarStyle used before Windows XP. 
	Thus 'Blocks' works only on Console Hosts.
	The ISE defaults to Windows animated Green with Yellow flashes.
	#>
}

	# Show splash
	$config|Show-SplashScreen 
	Start-Sleep -Seconds 5
	$total = ($Steps|%{$_.Weight}|Measure-Object -Sum).Sum
	$progress = 0
	foreach ($step in $Steps) {
		Update-Splash -Message $step.Name -Weight $step.Weight
		# Simulate work
		$StepDelay = $Delay * ($step.Weight/100)
		Start-Sleep -Seconds $StepDelay
	}
	# Final state
	$SplashRunspace.HashTable.ProgressMessage = '{0} Ready!' -f $AppName
	$SplashRunspace.HashTable.ProgressValue   = 100
	Start-Sleep -Seconds 1
	Close-SplashScreen
}
$SSIcon = Get-Icon
Start-Demo -Delay $Delay
}