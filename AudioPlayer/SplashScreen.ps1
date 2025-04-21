<#
.NOTES
-------------------------------------
Name:	SplashScreen.ps1
Version: 3.0 - 10/22/2022
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This script launches a Winform Splash Screen on a Background Thread.
There are 2 functions, Show-SplashScreen & Close-SplashScreen. This version uses
a Class to both define & create a SplashRunspace object used by the functions.
As a cmdlet you may run a demo of a Splash Screen but when imported as a
module you may call the Show-SplashScreen & Close-SplashScreen functions directly.
----------------------------------------------------------------------------------------

.DESCRIPTION
There are 2 primary functions, Show-SplashScreen & Close-SplashScreen, When calling 
Show-SplashScreen 5 parameters allow specifing the Calling Application Name, An Image 
to Display, and Optionally the form background color, the PictureBox background color
& a switch which causes the image to be used as a background image. By default the 
image is displayed inside a PictureBox Control.
---------------------------------------------------------------------------------------- 
.Parameter Demo
An optional switch that will cause a demo splash screen to be displayed.
.Parameter Delay
An optional value in seconds that may be used to alter the default 30 second demo duration.
.Parameter ImageSize
Optional sets the square image size for the demo.
.Parameter PbSizeMode
Optional sets the [Windows.Forms.PictureBoxSizeMode] for the demo.
.Parameter BgImage 
Optional switch causes the image to be background image
.Parameter BGILayout
Optional BackgroundImageLayout.  
#>
[CmdletBinding()]
param(
	[Parameter()][Switch]$Demo,
	[Parameter()][Int]$Delay=30,
	[Parameter()][Int]$ImageSize=256,
	[Parameter()][Windows.Forms.PictureBoxSizeMode]$PbSizeMode=[Windows.Forms.PictureBoxSizeMode]::Zoom,
	[Parameter()][Windows.Forms.ImageLayout]$BGILayout=[Windows.Forms.ImageLayout]::Zoom,
	[Parameter()][Switch]$BgImage)

<#
Declare the Custom Runspace Object used by Show-SplashScreen & Close-SplashScreen as
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
splash screen to close and the thread to be released. HashTable.SplashIsLoaded is an integer value when this value 
is zero the form initialization occurs, once initialization has occured this is incremented to 1 to prevent the form
being redefined in subsequent passes of a do-while loop. Once incremented the splash screen is displayed via the
ShowDialog() method. The splash screen remains active until the Close-SplashScreen function signals it to close
by setting HashTable.Flag to false. The second top-level property is Powershell a [Powershell] type object
representing the separate thread where the splash screen script will run. This property inherits all the attributes
of the [Powershell] object it contains. The last top-level property is Handle, this is the [IAsyncResult] object
returned when the background thread is fired by a call to the Splashrunspace.Powershell.BeginInvoke() method. This
objects properties provide status information about the background thread and is passed back when Close-SplashScreen 
is invoked via a SplashRunspace.PowerShell.EndInvoke() after the splash screen has closed and the Handle.IsCopleted 
property is true. The SplashRunspace.PowerShell property is disposed and the $Script:SplashRunspace variable is
reset to $Null. When you create a New SplashRunspace object you must pass the Name of the SplashRunspace.HashTable
in the associated splash screen script.
#>
Class SplashRunspace{
	#Properties
	[HashTable]$HashTable
	[PowerShell]$Powershell
	[IAsyncResult]$Handle

	#New Method Constructor
	SplashRunspace([String]$CommunicationsHashTableName){
		$This.HashTable = [HashTable]::Synchronized(@{})
		$This.HashTable.Flag = $True #Display Splash while True
		$This.HashTable.SplashIsLoaded = 0 #Value determines if Splash is loaded 0/1 = Off/On 
		$This.Powershell = [PowerShell]::Create()
		$This.Powershell.Runspace = [RunspaceFactory]::CreateRunspace()
		$This.Powershell.Runspace.Open()
		#Pass HashTable used for interthread communications.
		$This.Powershell.Runspace.SessionStateProxy.SetVariable($CommunicationsHashTableName,$This.HashTable)
	}
}
#endregion

#region Functions
<#
.NOTES
-------------------------------------
Name:	Function Show-SplashScreen
Version: 1.2 - 07/22/2022
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This function Displays a Splash Screen on a Background Thread.
----------------------------------------------------------------------------------------
.DESCRIPTION
This function has 6 parameters allowing specifing the Calling Application Name,
An Image to Display, Optionally the form background color, picturebox background color,
BackgroundImageLayout to Apply & a switch which causes the image to be used as a 
background image. By default the image is displayed inside a PictureBox Control.
---------------------------------------------------------------------------------------- 
.PARAMETER AppName - Alias: Name
Optional, Name of the application to be displayed - Defaults to "Application".

.PARAMETER Image - Alias: Img
Required, A Drawing.Image to display on the splash screen.

.PARAMETER FormBackColor - Alias: FBC
Optional form background color

.PARAMETER PicBackColor - Alias: PBC
Optional PictureBox background color

.PARAMETER PicBackColor - Alias: PBS
Optional PictureBox image size mode

.PARAMETER BGILayout - Alias: BGL
Optional, BackgroundImageLayout to Apply (Default: Zoom).

.PARAMETER BgImage - Alias: BGI
Optional, Switch - If present the image is displayed as a background image, 
otherwise the image is displayed in a PictureBox.

.EXAMPLE
Show-SplashScreen -AppName <Name> -Image <Image>
Displays the Splash Screen using a PictureBox
#>
function Show-SplashScreen{
param(
	[Parameter(Mandatory)][Alias('Img')][Drawing.Image]$Image,
	[Parameter()][Alias('Name')][String]$AppName = 'Application',
	[Parameter()][Alias('FBC')][Drawing.Color]$FormBackColor=[Drawing.SystemColors]::Control,
	[Parameter()][Alias('PBC')][Drawing.Color]$PicBackColor=[Drawing.SystemColors]::Control,
	[Parameter()][Alias('PBS')][Windows.Forms.PictureBoxSizeMode]$PbSizeMode=[Windows.Forms.PictureBoxSizeMode]::Zoom,
	[Parameter()][Alias('BGL')][Windows.Forms.ImageLayout]$BGILayout=[Windows.Forms.ImageLayout]::Zoom,
	[Parameter()][Alias('BGI')][Switch]$BgImage)

if($Null -ne $SplashRunspace){return} #Abort Call Already Active

$SplashScript = {
param(
	[Parameter(Mandatory)][Alias('Img')][Drawing.Image]$Image,
	[Parameter()][Alias('Name')][String]$AppName = 'Application',
	[Parameter()][Alias('FBC')][Drawing.Color]$FormBackColor=[Drawing.SystemColors]::Control,
	[Parameter()][Alias('PBC')][Drawing.Color]$PicBackColor=[Drawing.SystemColors]::Control,
	[Parameter()][Alias('PBS')][Windows.Forms.PictureBoxSizeMode]$PbSizeMode=[Windows.Forms.PictureBoxSizeMode]::Zoom,	
	[Parameter()][Alias('BGL')][Windows.Forms.ImageLayout]$BGILayout=[Windows.Forms.ImageLayout]::Zoom,
	[Parameter()][Alias('BGI')][Switch]$BgImage)

while($SyncHash.Flag){
	if($SyncHash.SplashIsLoaded -eq 0) # First Cycle of While
		{
		Add-Type -AssemblyName Windows.Forms

		$Splash = [Windows.Forms.Form]::New()
		$Timer = [Windows.Forms.Timer]::New()
		$ProgressBar = [Windows.Forms.ProgressBar]::New()
		$PSCoreFont = [Drawing.Font]::New('Segoe UI',9,[Drawing.FontStyle]::Regular)

		$Splash.Size = [Drawing.Size]::New(475,325)
		$Splash.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
		$Splash.Font = [Drawing.Font]::New('Arial', 12,[Drawing.FontStyle]::Regular)
		$Splash.BackColor = $FormBackColor
		$Splash.FormBorderStyle = [Windows.Forms.BorderStyle]::None
		$Splash.ControlBox = `
		$Splash.ShowInTaskbar = $False
		$SharedWidth = $Splash.Width - 25

		$Timer.Interval = 500
		$Timer_Tick = {
			$Timer.Enabled = $False
			if($SyncHash.Flag -eq $False){$Splash.Close()}
			if($ProgressBar.Value -eq $ProgressBar.Maximum){$ProgressBar.Value = $ProgressBar.Minimum}
			$ProgressBar.Value++
			$ProgressBar.Refresh()
			[Windows.Forms.Application]::DoEvents()
			$Timer.Enabled = $True
		}
		$Timer.Add_Tick($Timer_Tick)
		$Timer.Enabled = $True
		$Timer.Start()

		$ProgressBar.Style = [Windows.Forms.ProgressBarStyle]::Continuous
		$ProgressBar.Parent = $Splash
		$ProgressBar.Minimum = 0
		$ProgressBar.Maximum = 100
		$ProgressBar.Enabled = `
		$ProgressBar.Visible = $True
		$ProgressBar.Size = [Drawing.Size]::New($SharedWidth,20)
		$ProgressBar.Location = [Drawing.Point]::New(13,$Splash.Height-($ProgressBar.Height+10))
		$ProgressBar.BackColor = [Drawing.SystemColors]::Control
	
		$LblMsg = [Windows.Forms.Label]::New()
		$LblMsg.Parent = $Splash
		$LblMsg.Size = [Drawing.Size]::New($SharedWidth,20)
		$LblMsg.Location = [Drawing.Point]::New(13,$ProgressBar.Top-($LblMsg.Height+5))
		$LblMsg.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
		$LblMsg.BackColor = [Drawing.SystemColors]::Control
		$LblMsg.Font = $PSCoreFont
		$LblMsg.Text = "{0} Loading, Please Wait ..." -f $AppName
		$LblMsg.Visible = $True

		if($BgImage.IsPresent){
			$Splash.BackgroundImage = $Image
			$Splash.BackgroundImageLayout = $BGILayout
		}else{
			$PicBox = [Windows.Forms.PictureBox]::New()
			$PicBox.Parent = $Splash
			$PicBox.Size = [Drawing.Size]::New($SharedWidth,$LblMsg.Top-18)
			$PicBox.Location = [Drawing.Point]::New(13,13)
			$PicBox.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
			$PicBox.Image = $Image
			$PicBox.Visible = $True
			$PicBox.BackColor = $PicBackColor
			$PicBox.SizeMode = $PbSizeMode
		}

		#region Screen Drag
		$MouseDown = [Drawing.Point]::New(0,0)
		$DragPoint = [Drawing.Point]::New(0,0)

		$Splash_Shown = {
			$Splash.Activate()
			$Splash.TopMost=$True
			$Splash.BringToFront()
			$Splash.TopMost=$False}

		$Splash_MouseDown = {
			if($_.Button -eq [Windows.Forms.MouseButtons]::Left){$MouseDown = $_.Location}
		}

		$Splash_MouseMove = {
			if($_.Button -eq [Windows.Forms.MouseButtons]::Left){ 
				$DragPoint.X = $Splash.Location.X + ($_.X - $MouseDown.X)
				$DragPoint.Y = $Splash.Location.Y + ($_.Y - $MouseDown.Y)
				$Splash.Location = $DragPoint}
		}
		#endregion

		$InitialFormWindowState = $Splash.WindowState
		$Splash.Add_Load({$This.WindowState = $InitialFormWindowState})
		$Splash.Add_Shown($Splash_Shown)
		$Splash.Add_MouseDown($Splash_MouseDown)
		$Splash.Add_MouseMove($Splash_MouseMove)
		$SyncHash.SplashIsLoaded++
		$Splash.ShowDialog()
		}
	}
}

$Script:SplashRunspace = [SplashRunspace]::New('SyncHash')
#Pass Script & Parameters to runspace
[void]$SplashRunspace.Powershell.AddScript($SplashScript)
[void]$SplashRunspace.Powershell.AddParameter('AppName',$AppName)
[void]$SplashRunspace.Powershell.AddParameter('Image',$Image)
[void]$SplashRunspace.Powershell.AddParameter('FormBackColor',$FormBackColor)
[void]$SplashRunspace.Powershell.AddParameter('PicBackColor',$PicBackColor)
[void]$SplashRunspace.Powershell.AddParameter('PbSizeMode',$PbSizeMode)
if($BgImage.IsPresent)
	{
	[void]$SplashRunspace.Powershell.AddParameter('BgImage')
	[void]$SplashRunspace.Powershell.AddParameter('BgiLayout',$BGILayout)
	}
$SplashRunspace.Handle = $SplashRunspace.PowerShell.BeginInvoke()
}

<#
.NOTES
-------------------------------------
Name:	Function Close-SplashScreen
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
function Close-SplashScreen{
$SplashRunspace.HashTable.Flag = $False	#Close Screen
if($SplashRunspace.Handle.IsCompleted)
	{#Dispose of Thread
	$SplashRunspace.PowerShell.EndInvoke($SplashRunspace.Handle)
	$SplashRunspace.PowerShell.Dispose()
	$Script:SplashRunspace = $Null
	}
}
#endregion

if($Demo.IsPresent){
	#region Add Custom DLL - Data Type
	<#BlueflameDynamics.IconTools Class#>
	Import-Module -Name .\Exists.ps1 -Force
	$DLLPath = '.\BlueflameDynamics.IconTools.dll'
	if((Test-Exists -Mode File -Location $DLLPath) -eq $True){
		Add-Type -Path $DLLPath
		Start-Sleep -Seconds 1
		$SSImage = [BlueflameDynamics.IconTools]::ExtractIcon(
			$Env:WinDir+'\System32\ImageRes.dll',311,$ImageSize)
	}
	# Default Image (Powershell Icon) used for Demo/Testing 
	#endregion
	Show-SplashScreen -Image $SSImage -AppName 'SplashScreen Demo' -FBC ([Drawing.Color]::CornflowerBlue) -BGI:$BgImage -BGL $BGILayout
	Start-Sleep -Seconds $Delay
	Close-SplashScreen
}