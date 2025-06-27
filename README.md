# Powershell Utilities

My name is Randy Turner, I'm a retired Developer\Programmer with 50 years of experience in standalone & network applications development, and network administration.

The purpose of this repository is to share many of my most popular Powershell Utilities and Scripts
created since Powershell was first released.

This repository includes a collection of Powershell Utilities and Scripts many of which may be used
as reusable modules via the Import-Module command in building more complex scripts. One such script is
the 'WinFormsLibrary.ps1' which includes functions for many common Windows WinForm dialogs.

The repository contains folders for each application.

I've also included both a 32-bit & 64-bit version of my PSRun3 Windows Powershell Host application, which provides a means of running Windows Powershell sceipts without a console window. For command syntax run: 
PSRun3.exe -help

## File Catalog:
 [AudioPlayer Applicationn](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/AudioPlayer)
 | Filename     | Description  |
 | ------------ | ------------ |                                       
 | AppRegistry.ps1 | Application Registry Functions | 
 | AudioPlayer.ps1 | AudioPlayer Main Module | 
 | AudioPlayerEnums.ps1 | Enumerated Values | 
 | AudioPlayerIcons.dll | Icon Library |  
 | BlueflameDynamics.IconTools.dll | Icon Utils Library | 
 | Class_IconCatalogItem.ps1 | Class for Icon Access | 
 | Exists.ps1 | .Net Directory/File existance test function | 
 | ListviewSearchLib.ps1 | ListLiew Search Library | 
 | ListViewSortLib.ps1 | ListView Sort Library | 
 | PCVolumeControl.ps1 | Wrapper for C# code calling the Windows APIs for PC Audio Volume Control on Windows Vista & above. | 
 | PropertySheetDialog.ps1 | Custom Property Sheet Dialog | 
 | PS_Audio_Player_Help.txt | AudioPlayer Help File | 
 | PSRun3.exe | 32-bit Powershell Host | 
 | PSRun3X64.exe | 64-bit Powershell Host | 
 | SplashScreen.ps1 | WinForm Splash Screen Functions | 
 | UtilitiesLib.ps1 | Utility Function Library | 
 | WinFormsLibrary.ps1 | WinForms Dialog\Utility Function Library | 

 [Build-HexDump Cmdlet](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Build-HexDump)
 | Filename     | Description  |
 | ------------ | ------------ |                                        
 | Build-HexDump.ps1 | Hexidecimal File Dump Cmdlet | 
                        
 [Convert-Image Cmdlet](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Convert-Image)
 | Filename     | Description  |
 | ------------ | ------------ |                                          
 | Convert-Image.ps1 | Functions for resizing & converting an image file type. | 
                        
 [Convert-SpacesToTabs Cmdlet](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Convert-SpacesToTabs)
 | Filename     | Description  |
 | ------------ | ------------ |                                   
 | Convert-SpacesToTabs.ps1 | Cmdlet to Convert Spaces to Tabs |
          
 [Create-PSListing Cmdlet Group](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Create-PSListing)
 | Filename     | Description  |
 | ------------ | ------------ |                                       
 | BlueflameDynamics.IconTools.dll | Icon Utils Library |
 | Create-PSListing.ps1 | Powershell Script Listing Cmdlet |
 | Create-PSListing.SampleListing.pdf | Sample Powershell Script Listing |
 | Create-PSListingGUI.json | Powershell Script Listing GUI configuration json | 
 | Create-PSListingGUI.ps1 | Powershell Listing GUI front end |
 | Create-PSListingGUI.SampleListing.pdf | Sanple Script Listing | 
 | Create-PsListingGUI_Form.png | Powershell Script Listing GUI Form Sample |
 | Exists.ps1 | .Net Directory/File existance test function | 
 | OSInfoLib.ps1 | Windows OS Info Library | 
 | PSRun3.exe | 32-bit Powershell Host | 
 | PSRun3X64.exe | 64-bit Powershell Host |  
 | WinFormsLibrary.ps1 | WinForms Function Library | 

 [Export-MediaFileDetailsEx Cmdlet Group](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Export-MediaFileDetailsEx)
 | Filename     | Description  |
 | ------------ | ------------ | 
 | AES-Email.ps1 | AES Email Function Library |
 | Exists.ps1 | .Net Directory/File existance test function |                           
 | Export-MediaFileDetailsEx.ps1 | Cmdlet for Export of Media File Extended Properties |
 | Get-NetMediaInfo.ps1 | Gets Network Media File Extended Properties (requires configration changes)
 | MetadataIndexLib.ps1 | MetaData Index Access Functions |
 | ScriptPathInfo.ps1 | Script Information Library |
 | Suspend-PowerPlan.ps1 | Function Library to Suspend the system Power Plan for long running Tasks |
 | Test-InternetConnection.ps1 | Function Library to Test for a live internet connection |

 [MovieLauncher Application](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/MovieLauncher)
 | Filename     | Description  |
 | ------------ | ------------ |                                          
 | BlueflameDynamics.IconTools.dll | Icon Utils Library | 
 | Class_IconCatalogItem.ps1 | Class for Icon Access |
 | Exists.ps1 | .Net Directory/File existance test function |
 | GetSplitPathLib.ps1 | File Path Split function library |
 | Invoke-CopyFile.ps1 | Cmdlet\Library to copy a file and display a Progress Window. |
 | ListviewSearchLib.ps1 | ListLiew Search Library | 
 | ListViewSortLib.ps1 | ListView Sort Library |
 | MetadataIndexLib.ps1 | MetaData Index Access Functions |
 | Movie Launcher.lnk | Sample Windows Shortcut File |
 | MovieLauncher.ps1 | Main Movie Launcher Cmdlet |
 | Music Videos.lnk | Sample Windows Shortcut File |
 | PSMovieLauncherIcons.dll | Icon Library |
 | PSRun3.exe | 32-bit Powershell Host | 
 | PSRun3X64.exe | 64-bit Powershell Host |  
 | Recorded TV.lnk | Sample Windows Shortcut File |
 | UtilitiesLib.ps1 | Utility Function Library | 
 | WinFormsLibrary.ps1 | WinForms Dialog\Utility Function Library |

 [PSRun-Host-Application (No Console Window)](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/PSRun-Host-Application)
 | Filename     | Description  |
 | ------------ | ------------ |                                
 | PSRun3.exe | 32-bit Powershell Host | 
 | PSRun3_About.png | Sample About Box |
 | PSRun3_Help.png | Sample Help |
 | PSRun3X64.exe | 64-bit Powershell Host | 
                                              
 [Set-OpticalDriveState Cmdlet](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Set-OpticalDriveState)
 | Filename     | Description  |
 | ------------ | ------------ |                                  
 | Set-OpticalDriveState.ps1 | Cmdlet to set\get Optical Drive State |
        
 [Show-ColorPicker\Get-ColorInfo Cmdlets](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Show-ColorPicker)
 | Filename     | Description  |
 | ------------ | ------------ |                                       
 | BlueflameDynamics.IconTools.dll | Icon Utils Library |
 | Exists.ps1 | .Net Directory/File existance test function |
 | Get-ColorInfo.ps1 | Cmdlet to Get Window Color Info | 
 | Get-ColorInfo_Desc.pdf | Get-ColorInfo.ps1 Description Document |
 | PSColorPicker.png | Sample ColorPicker |
 | Show-ColorPicker.ps1 | Cmdlet to display ColorPicker |
 | Show-ColorPicker_Desc.pdf | Show-ColorPicker.ps1 Description Document |
 | WinFormsLibrary.ps1 | WinForms Dialog\Utility Function Library |
                   
 |[Show-DynamicMenu Cmdlet](https://github.com/BlueflameDynamics/Powershell-Utilities/tree/main/Show-DynamicMenu)
 | Filename     | Description  |
 | ------------ | ------------ |                                       
 | Show-DynamicMenu.ps1 | Cmdlet\Library to display a Dynamic Menu |                 



