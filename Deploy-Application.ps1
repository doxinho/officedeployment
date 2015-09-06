<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
	Deploy-Application.ps1
.EXAMPLE
	Deploy-Application.ps1 -DeployMode 'Silent'
.EXAMPLE
	Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.EXAMPLE
	Deploy-Application.ps1 -DeploymentType Uninstall
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = ''
	[string]$appName = 'Microsoft Office 365'
	[string]$appVersion = ''
	[string]$appArch = ''
	[string]$appLang = ''
	[string]$appRevision = ''
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '09/05/2015'
	[string]$appScriptAuthor = ''
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.5'
	[string]$deployAppScriptDate = '08/17/2015'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	# Office Variables
	 [string[]] $dirOffice = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
	 [string[]] $dirOffice32 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office"
	 [string[]] $dirOfficeC2R = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office 15"
	 [string[]] $architecture = (Get-WmiObject win32_processor | Where-Object{$_.deviceID -eq "CPU0"}).AddressWidth
	 [string[]] $officeExecutables = 'excel.exe', 'groove.exe', 'infopath.exe', 'onenote.exe', 'outlook.exe', 'mspub.exe', 'powerpnt.exe', 'winword.exe', 'winproj.exe', 'visio.exe'	

	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
		
		# Show installation prompt
		 Show-InstallationPrompt -Title "Office Upgrade" -Message "You will be removing your existing Office isntallation and upgrading to Microsoft Office 365. Please click below to begin the installation." -ButtonMiddleText "OK" -Icon Exclamation -PersistPrompt -MinimizeWindows $true

		# Show Welcome Message, close apps, prompt to save, check disk space
		 Show-InstallationWelcome -CloseApps "iexplore,PWConsole,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,communicator,lync" -PromptToSave -ForceCloseAppsCountdown 120 -BlockExecution -CheckDiskSpace
		 
		# Display Pre-Install cleanup status
		 Show-InstallationProgress "Uninstalling previous versions of Microsoft Office. This may take some time. Please wait..."
			
		# Remove any previous version of Office (if required)

		# Office 2003 (32-bit systems)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice32 -ChildPath "Office11\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2003 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path 'cscript.exe' -Parameters "`"$dirFiles\OffScrub03.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2003
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office11\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2003 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path 'cscript.exe' -Parameters "`"$dirFiles\OffScrub03.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2007 (32-bit)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office12\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2007 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path 'cscript.exe' -Parameters "`"$dirFiles\OffScrub07.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2010 (32-bit)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office14\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2010 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path "cscript.exe" -Parameters "`"$dirFiles\OffScrub10.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2010 (64-bit)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice32 -ChildPath "Office14\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2010 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path "cscript.exe" -Parameters "`"$dirFiles\OffScrub10.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2013 (32-bit)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office15\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2013 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path "cscript.exe" -Parameters "`"$dirFiles\OffScrub13.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2013 (64-bit)
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOffice32 -ChildPath "Office15\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2013 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path "cscript.exe" -Parameters "`"$dirFiles\OffScrub13.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
				Break
			}
		}

		# Office 2013/365 C2R (32-bit & 64-bit); note ExitCode 34 ignored to due failures removing 365 without it
		 ForEach ($officeExecutable in $officeExecutables) {
			If (Test-Path -Path (Join-Path -Path $dirOfficeC2R -ChildPath "root\Office15\$officeExecutable") -PathType Leaf) {
				Write-Log -Message 'Microsoft Office 2013 C2R was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
				Execute-Process -Path "cscript.exe" -Parameters "`"$dirFiles\OffScrub_O15c2r.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,34'
				Break
			}
		}

		# Remove Microsoft SkyDrive
		 Remove-MSIApplications "Microsoft SkyDrive"

		# Remove Microsoft Guide to The Ribbion
		 Remove-MSIApplications "Microsoft Guide to the Ribbon"
		 
		# Remove Microsoft Office 2007 Help Tab
		 Remove-MSIApplications "Microsoft Office 2007 Help Tab"
		 
		# Remove Microsoft Conferencing Add-in for Microsoft Office Outlook
		 Remove-MSIApplications "Microsoft Conferencing Add-in for Microsoft Office Outlook"
		 
		# Remove Microsoft Office Live Meeting 2007
		 Remove-MSIApplications "Microsoft Office Live Meeting 2007"
		 
		# Remove Microsoft Office 2010 Interactive Guide
		 Remove-MSIApplications "Microsoft Office 2010 Interactive Guide"
		 
		# Remove Microsoft Office 2010 User Resources
		 Remove-MSIApplications "Office 2010 User Resources"
		 
		# Remove Microsoft Office Communicator 2007
		 Remove-MSIApplications "Microsoft Office Communicator 2007"

		# Remove Microsoft Office 2007 Primary Interop Assemblies
		 Remove-MSIApplications "Microsoft Office 2007 Primary Interop Assemblies"

		# Remove Microsoft Office Access database engine 2007 (English)
		 Remove-MSIApplications "Microsoft Office Access database engine 2007 (English)"

		# Remove Microsoft Office File Validation Add-in
		 Remove-MSIApplications "Microsoft Office File Validation Add-in"

		# Remove Microsoft Office Suite Activation Assistant
		 Remove-MSIApplications "Microsoft Office Suite Activation Assistant"
		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		# Installing Office 365 Pro Plus
		 Show-InstallationProgress "Installing Office 365. This may take some time. Please wait..."
		 Execute-Process -FilePath "$dirFiles\Office365\setup.exe" -Parameters "/configure `"$dirFiles\Office365\Installation.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		# Suppress First Run Dialogs
		 If ($architecture -eq 64) {
		     & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings" /v Count /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\FirstRun" /v BootedRTM /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\FirstRun" /v disablemovie /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\General" /v shownfirstrunoptin /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\General" /v ShownFileFmtPrompt /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\PTWatson" /v PTWOptIn /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common" /v qmenable /t REG_DWORD /d 1 /f
		 } else {
		     & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings" /v Count /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\FirstRun" /v BootedRTM /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\FirstRun" /v disablemovie /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\General" /v shownfirstrunoptin /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\General" /v ShownFileFmtPrompt /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common\PTWatson" /v PTWOptIn /t REG_DWORD /d 1 /f
			 & reg add "HKLM\SOFTWARE\Microsoft\Office\15.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\15.0\Common" /v qmenable /t REG_DWORD /d 1 /f
		 }

		# Show Dialog Box
		 Show-InstallationProgress "Installation complete."
		 Show-DialogBox -Title "Installation complete" -Text "The installation is complete. You will now be prompted to restart your computer." -Icon Information
		
		# Prompt for a restart (if running as a user, not installing components and not running on a server)
		If ((-not $addComponentsOnly) -and ($deployMode -eq 'Interactive') -and (-not $IsServerOS)) {
			Show-InstallationRestartPrompt
		}
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}