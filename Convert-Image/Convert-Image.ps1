Function Convert-Image
{
<#
.NOTES
	Author: Randy Turner
	Version: 1.0a - 04/13/2020
		
.SYNOPSIS
	Resize and optionally convert an image file.

.DESCRIPTION
	The Convert-Image cmdlet will set a new size, resolution, and\or convert an image file.
	----------------------------------------------------------------------------------------
	Security Note: This is an unsigned script, Powershell security may require you run the
	Unblock-File cmdlet with the Fully qualified filename before you can run this script,
	assuming PowerShell security is set to RemoteSigned.
	----------------------------------------------------------------------------------------

.PARAMETER ImgFile
	Specifies an input image file. 

.PARAMETER Destination
	Specifies a destination for resized file(s). Default is current location (Get-Location).

.PARAMETER SaveAs
	Specifies an output image type allowing type conversion.

.PARAMETER WidthPx
	Specifies a width of image in pixels. 

.PARAMETER HeightPx
	Specifies a height of image in pixels.

.PARAMETER DPIWidth
	Specifies a vertical resolution. 

.PARAMETER DPIHeight
	Specifies a horizontal resolution.

.PARAMETER Overwrite
	Specifies a destination exist then overwrite it without prompt. 

.PARAMETER FixedSize
	Set to a fixed size and do not try to scale the aspect ratio.

.PARAMETER DetectLandscape
	Swap Height and Width when Landscape source image is detected. 

.PARAMETER RemoveSource
	Remove source file after conversion. 

.EXAMPLE
	PS C:\> Get-ChildItem 'H:\Test\In\*.jpg' | Sort-Object -Property FullName | Convert-Image -Destination "H:\Test\Out" -WidthPx 800 -HeightPx 800 -DetectLandscape -SaveAs Png -Verbose
	VERBOSE: Image 'H:\Test\In\1.jpg' was resized from: 600x598 to 800x797 and saved as: 'H:\Test\Out\1.png'
	VERBOSE: Image 'H:\Test\In\2.jpg' was resized from: 600x598 to 800x797 and saved as: 'H:\Test\Out\2.png'
	VERBOSE: Image 'H:\Test\In\3.jpg' was resized from: 600x598 to 800x797 and saved as: 'H:\Test\Out\3.png'
#>
[CmdletBinding(
	SupportsShouldProcess=$True,
	ConfirmImpact="Low"
)]
Param
(
	[Parameter(Mandatory,
		ValueFromPipeline,
		ValueFromPipelineByPropertyName)]
	[Alias("Image")]
	[String[]]$FullName,
	[Parameter(Mandatory = $False)][Alias('Save')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Bmp','Gif','Jpeg','Png','Tiff')]
	[String]$SaveAs='',
	[String]$Destination = $(Get-Location),
	[Int]$WidthPx,
	[Int]$HeightPx,
	[Int]$DPIWidth,
	[Int]$DPIHeight,
	[Switch]$Overwrite,
	[Switch]$DetectLandscape,
	[Switch]$FixedSize,
	[Switch]$RemoveSource,
	[Switch]$DetailedReport
)

Begin
{
	$ImgTypes = @{Bmp = '.bmp';Gif = '.gif';Jpeg = '.jpg';Png = '.png';Tiff = '.tif'}
	$ImgGUIDs = @{
		'B96B3CAB-0728-11D3-9D7B-0000F81EF32E' = 'Bmp';
		'B96B3CB0-0728-11D3-9D7B-0000F81EF32E' = 'Gif';
		'B96B3CAE-0728-11D3-9D7B-0000F81EF32E' = 'Jpeg';
		'B96B3CAF-0728-11D3-9D7B-0000F81EF32E' = 'Png';
		'B96B3CB1-0728-11D3-9D7B-0000F81EF32E' = 'Tiff'}
}

Process
{
	Foreach($ImageFile in $FullName)
	{
		If(Test-Path -Path $ImageFile)
		{
			$OldImage = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $ImageFile
			$OldWidth = $OldImage.Width
			$OldHeight = $OldImage.Height
			$H = $HeightPx
			$W = $WidthPx
			If($OldImage.Width -gt $OldImage.Height -and $DetectLandscape.IsPresent)
				{
				#Landscape Image Detected
				$H = $WidthPx
				$W = $HeightPx
				}

			If($WidthPx -eq $Null){$WidthPx = $OldWidth}
			If($HeightPx -eq $Null){$HeightPx = $OldHeight}

			If($FixedSize.IsPresent)
			{
				$NewWidth = $WidthPx
				$NewHeight = $HeightPx
			}
			else
			{
				If($OldWidth -lt $OldHeight)
				{
					$NewWidth = $W
					[Int]$NewHeight = [Math]::Round(($NewWidth*$OldHeight)/$OldWidth)
						
					If($NewHeight -gt $HeightPx)
					{
						$NewHeight = $H
						[Int]$NewWidth = [Math]::Round(($NewHeight*$OldWidth)/$OldHeight)
					}
				}
				else
				{
					$NewHeight = $H
					[Int]$NewWidth = [Math]::Round(($NewHeight*$OldWidth)/$OldHeight)
						
					If($NewWidth -gt $WidthPx)
					{
						$NewWidth = $W
						[Int]$NewHeight = [Math]::Round(($NewWidth*$OldHeight)/$OldWidth)
					}
				}
			}

			$ImageProperty = Get-ItemProperty -Path $ImageFile
			$SaveLocation = Join-Path -Path $Destination -ChildPath ($ImageProperty.Name)
			$NewImage = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $NewWidth,$NewHeight

			$Graphics = [System.Drawing.Graphics]::FromImage($NewImage)
			$Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
			$Graphics.DrawImage($OldImage, 0, 0, $NewWidth, $NewHeight) 

			$ImageFormat = $OldImage.RawFormat
			$ImgFormatStr = $ImageFormat.Guid.ToString().ToUpper()
			$ImgType = $ImgGUIDs.Item($ImgFormatStr)
			If($SaveAs.Length -ne 0 -and $SaveAs -ne $ImgType)
				{
				# Image type conversion requested!
				$SaveLocation = -join($SaveLocation.Substring(0,$SaveLocation.LastIndexOf('.')),$ImgTypes.Item($SaveAs))
				$ImageFormat = [System.Drawing.Imaging.ImageFormat]::$SaveAs 
				}
			$OldImage.Dispose()
			If($DPIWidth -and $DPIHeight)
			{
				$NewImage.SetResolution($DPIWidth,$DPIHeight)
			}

			If(!$Overwrite.IsPresent)
			{
				If(Test-Path -Path $SaveLocation)
				{
					$Title = "A file already exists: $SaveLocation"
					$ChoiceOverwrite = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList "&Overwrite"
					$ChoiceCancel = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList "&Cancel"
					$Options = [System.Management.Automation.Host.ChoiceDescription[]]($ChoiceCancel, $ChoiceOverwrite)	

					If(($host.ui.PromptForChoice($Title, $null, $Options, 1)) -eq 0)
					{
						Write-Verbose -Message "Image '$ImageFile' exists in destination location - skipped"
						Continue
					}
				}
			}

			$NewImage.Save($SaveLocation,$ImageFormat)
			$NewImage.Dispose()
				
			If($DetailedReport.IsPresent)
				{
				Write-Verbose -Message "Image '$ImageFile' was resized from: $($OldWidth)x$($OldHeight) to $($NewWidth)x$($NewHeight) and saved as: '$SaveLocation'"
				}
			Else
				{
				$FN = Split-Path -Path $ImageFile -Leaf
				Write-Verbose -Message "Image '$FN' was resized from: $($OldWidth)x$($OldHeight) to $($NewWidth)x$($NewHeight)'"
				}

			If($RemoveSource.IsPresent)
			{
				Remove-Item -Path $ImageFile -Force
				Write-Verbose -Message "Input Image: '$ImageFile', was removed"
			}
		}
	}
}

End{}
}