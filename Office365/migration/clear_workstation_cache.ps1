<#
    The purpose of this script is to clear out all O365 related credentials on a workstation when migrating from one Office 365 tenant to another Office 365 tenant but maintaining the same userprincipalname. If you do not clear out credentials, the cached credentials will attempt to login to the old tenant once the migration is complete. You can uncomment the final line of the script to display a message to the user. 
    
    This script is only designed to run once if it is successful, it will generate a file at $outputlocation\full-script-complete.txt once all steps of the script have completed. If you wish to re-run the script, delete $outputlocation\full-script-complete.txt otherwise the script steps will just skip. This file is only generated at the end of the script, so if the script is interrupted in the middle, the file will not be generated
#>


$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
# Add a counter to know when to skip
$counter=0
# Test the path $env:USERPROFILE + "\migration is there, create it if it doesn't exist
$Migrationfolder = Test-Path $env:USERPROFILE"\migration"
If($Migrationfolder -eq $False) {
    Try {
        $ErrorActionPreference = 'stop'
        $null = New-Item -Path $env:USERPROFILE -Name "Migration" -ItemType "directory"
    }
    Catch {
        # If there is an error, log the error
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Host $error[0].Exception.GetType().FullName
	    Write-Host $errormessage -ForegroundColor Red
    }
}

# set log and output locations
$loglocation = $env:USERPROFILE + "\migration\log.txt"
$outputlocation = $env:USERPROFILE + "\migration"

# Begin Teams sign out
" " | out-file -append $loglocation
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 1/7): Teams sign out"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation

# Clear Teams cached folders under %appdata%\Microsoft\Teams
# Check if teams folder cache was already cleared by script
$TeamsCache = Test-Path -Path $outputlocation\teams-cache-cleared.txt
If ($TeamsCache -eq $false){
	# Get MS Teams process. Only using 'SilentlyContinue' as we test this below
	$TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

	# Get Outlook process. Only using 'SilentlyContinue' as we test this below
	$OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}

	If ($OutlookProcess) {
		# If 'Outlook' process is running, stop it else do nothing
		$OutlookProcess | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Outlook process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
    $TeamsFolderCheck = Test-Path -Path $env:APPDATA\"Microsoft\Teams"
    If ($TeamsFolderCheck -eq $true) {
        # Check if 'Teams' folder exists in %APPDATA%\Microsoft\Teams
        Try { 
            $ErrorActionPreference = 'stop'
            $Blob = Test-Path -Path $env:APPDATA\"Microsoft\teams\blob_storage"
            If( $Blob -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage" | Remove-Item -Recurse
            }
            $db = Test-Path -Path $env:APPDATA\"Microsoft\teams\databases"
            If( $db -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases" | Remove-Item -Recurse
            }
            $cache = Test-Path -Path $env:APPDATA\"Microsoft\teams\cache"
            If( $cache -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache" | Remove-Item -Recurse
            }
            $gpucache = Test-Path -Path $env:APPDATA\"Microsoft\teams\gpucache"
            If( $gpucache -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache" | Remove-Item -Recurse
            }
            $Indexeddb = Test-Path -Path $env:APPDATA\"Microsoft\teams\Indexeddb"
            If( $Indexeddb -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb" | Remove-Item -Recurse
            }
            $local = Test-Path -Path $env:APPDATA\"Microsoft\teams\Local Storage"
            If( $local -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage" | Remove-Item -Recurse
            }
            $tmp = Test-Path -Path $env:APPDATA\"Microsoft\teams\tmp"
            If( $tmp -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp" | Remove-Item -Recurse
            }
            $cookies = Test-Path -Path $env:APPDATA\"Microsoft\teams\Cookies"
            If( $cookies -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Cookies" | Remove-Item
            }
            $storage = Test-Path -Path $env:APPDATA\"Microsoft\teams\storage.json"
            If( $storage -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\storage.json" | Remove-Item
            }
			# If no errors renaming folder, log success
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams cache was cleared and user was signed out"
			#write-host $output
			$output | out-file -append $loglocation
			$null = New-Item $outputlocation\teams-cache-cleared.txt
			$counter ++
        }
        Catch {
			# If there is an error, log the error
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$errormessage=$timestamp + " ERROR: " + $_.ToString()
			#write-host $error[0].Exception.GetType().FullName
			#write-host $errormessage -ForegroundColor Red
			$errormessage | out-file -append $loglocation
        }
    }
    Else {
        # If 'Teams' folder does not exist notify user then break
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Teams folder not found"
	    #write-host $output
	    $output | out-file -append $loglocation
    }
}
Else {
    # If 'Teams' folder was already renamed then skip
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Teams cache was already cleared"
    #write-host $output
    $output | out-file -append $loglocation
}

# Check if HomeUserUpn for teams was cleared by script
$TeamsUserClear = Test-Path -Path $outputlocation\teams-homeuserupn-cleared.txt
If ($TeamsUserClear -eq $false){
	# Get MS Teams process. Only using 'SilentlyContinue' as we test this below
	$TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

	# Get Outlook process. Only using 'SilentlyContinue' as we test this below
	$OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}

	If ($OutlookProcess) {
		# If 'Outlook' process is running, stop it else do nothing
		$OutlookProcess | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Outlook process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
    Try { 
        $ErrorActionPreference = 'stop'
        Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Teams" -Name "HomeUserUpn"
        # If no errors clearing HomeUserUpn, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Registry entry deleted: HKCU:\Software\Microsoft\Office\Teams\HomeUserUpn"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\teams-homeuserupn-cleared.txt
		$counter ++
    }
    Catch {
		# If there is an error, log the error
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-host $error[0].Exception.GetType().FullName
		#write-host $errormessage -ForegroundColor Red
		$errormessage | out-file -append $loglocation
    }
}
# If HomeUserUpn was cleared before then skip
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Teams HomeUserUpn was already cleared "
    #write-host $output
    $output | out-file -append $loglocation
}
# Modify desktop-Config.Json under %appdata%\Microsoft\Teams
# Check if modify desktop-Config.json was already done by script
$TeamsModifyDesktopConfig = Test-Path -Path $outputlocation\teams-modify-desktop-config.txt
If($TeamsModifyDesktopConfig -eq $false) {
	# Get MS Teams process. Only using 'SilentlyContinue' as we test this below
	$TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

	# Get Outlook process. Only using 'SilentlyContinue' as we test this below
	$OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}

	If ($OutlookProcess) {
		# If 'Outlook' process is running, stop it else do nothing
		$OutlookProcess | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Outlook process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
    Try { 
		# Import desktop-Config.json
		$ErrorActionPreference = 'stop'
		$TeamsFolder = "$env:APPDATA\Microsoft\Teams"
		$SourceDesktopConfigFile = "$TeamsFolder\desktop-config.json"
		$desktopConfig = (Get-Content -Path $SourceDesktopConfigFile | ConvertFrom-Json)
		# If no errors importing desktop-config, log success
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " " + $TeamsFolder + "\desktop-Config/json imported successfuly"
		#write-host $output
		$output | out-file -append $loglocation
		# Modify desktop-Config.json
		if($desktopConfig.isLoggedOut -ne $null) {
			$desktopConfig.isLoggedOut = $true
		}
		if($desktopConfig.upnWindowUserUpn -ne $null) {
			$desktopConfig.upnWindowUserUpn =""; #The email used to sign in
		}
		if($desktopConfig.userUpn -ne $null) {
			$desktopConfig.userUpn ="";
		}
		if($desktopConfig.userOid -ne $null) {
			$desktopConfig.userOid ="";
		}
		if($desktopConfig.userTid -ne $null) {
			$desktopConfig.userTid = "";
		}
		if($desktopConfig.homeTenantId -ne $null) {
			$desktopConfig.homeTenantId ="";
		}
		if($desktopConfig.webAccountId -ne $null) {
			$desktopConfig.webAccountId="";
		}
		$desktopConfig | ConvertTo-Json -Compress | Set-Content -Path $SourceDesktopConfigFile -Force
		# If no errors modifying desktop-config, log success
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " " + $TeamsFolder + "\desktop-Config/json modified successfuly"
		#write-host $output
		$output | out-file -append $loglocation
		$null = New-Item $outputlocation\teams-modify-desktop-config.txt
		$counter ++
    }
    Catch {
		# If there is an error, log the error
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-host $error[0].Exception.GetType().FullName
		#write-host $errormessage -ForegroundColor Red
		$errormessage | out-file -append $loglocation
    }
}
# If script desktop-Config.json was already modified then skip
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: desktop-Config.json was already modified"
    #write-host $output
    $output | out-file -append $loglocation
}

# Begin Outlook new profile creation
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 2/7): New profile creation for Outlook"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation
# Check if new profile was created
$NewOutlookProfile = Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New
If ($NewOutlookProfile -eq $True) {
    # If profile was created, skip profile creation
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Outlook profile 'Outlook-New' was already created"
    #write-host $output
    $output | out-file -append $loglocation
}
Else {
    # If profile was was not created, create new profile
    # If 'Outlook' process is running, stop it else do nothing
    # Get Outlook process. Only using 'SilentlyContinue' as we test this below
    $OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue
    If ($OutlookProcess) {
        $OutlookProcess | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Outlook process was running, so we stopped it"
        #write-host $output
		$output | out-file -append $loglocation
    }
    Try {
        # Create new Outlook profile named 'Outlook-New'
        $ErrorActionPreference = 'stop'
        $null = New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New
        # If no errors creating the new profile, log success
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " RegKey created: HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New"
		#write-host $output
		$output | out-file -append $loglocation
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-warning $errormessage
		$errormessage | out-file -append $loglocation
    }
}
# Check if default profile has been set to Outlook-New
$DefaultProfile = Get-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook -Name DefaultProfile -ErrorAction SilentlyContinue
If($DefaultProfile.defaultprofile -eq "Outlook-New") {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Default Outlook profile already set to Outlook-New"
    #write-host $output
	$output | out-file -append $loglocation
}
Else {
    # If 'Outlook' process is running, stop it else do nothing
    # Get Outlook process. Only using 'SilentlyContinue' as we test this below
    $OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue
    If ($OutlookProcess) {
        $OutlookProcess | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
        #write-host "Outlook process was running, so we stopped it" -ForegroundColor Green
    }
    Try {
        # Set default Outlook profile to 'Outlook-New'
        $ErrorActionPreference = 'stop'
        $null = New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook -Name DefaultProfile -PropertyType String -Value "Outlook-New" -Force
        # If no errors setting the default Outlook profile, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Reg string created: 'HKCU:\Software\Microsoft\Office\16.0\Outlook\DefaultProfile' with value 'Outlook-New'"
        #write-host $output
	    $output | out-file -append $loglocation
		$counter ++
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

# Begin OneDrive sign out and unlink
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 3/7): OneDrive sign out and unlink"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation
# Clear OneDrive credentials unless script has been run previously
$OneDriveCache = Test-Path -Path $outputlocation\onedrive-cached-creds-cleared.txt
If ($OneDriveCache -eq $false){
    Try {
		$CheckNuget = Get-PackageProvider -listavailable nuget -ErrorAction SilentlyContinue
		If($CheckNuget -eq $null ) {
            $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -scope currentuser -force
        }
        $CheckPsCred = Get-Module -listavailable pscredentialmanager -ErrorAction SilentlyContinue
		If($CheckPsCred -eq $null ) {
            $null = Install-Module -Name pscredentialmanager -Scope CurrentUser -force
        }
        $CheckCredMan = Get-Module -listavailable CredentialManager -ErrorAction SilentlyContinue
		If($CheckCredMan -eq $null ) {
            $null = Install-Module -Name CredentialManager -Scope CurrentUser -force
        }
		$onedrive = Get-CachedCredential | where {$_.name -like "*onedrive*"}
		If($onedrive -ne $null) {
			foreach($cred in $onedrive) {
				remove-storedcredential -target $cred.name
				$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
				$output = $timestamp + " Cleared Cached OneDrive Credentials"
				#write-host $output
				$output | out-file -append $loglocation
			}
			$null = New-Item -Path $outputlocation\onedrive-cached-creds-cleared.txt
			$counter ++
			
		}
		Else {
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " No Cached OneDrive Credentials to clear"
			#write-host $output
			$output | out-file -append $loglocation
			$null = New-Item -Path $outputlocation\onedrive-cached-creds-cleared.txt
			$counter ++
		}
    }
	Catch {
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-warning $errormessage
		$errormessage | out-file -append $loglocation
	}
}

# Clear registry path HKCU:\Software\Microsoft\OneDrive\Accounts\* to unlink OneDrive account 
$OneDriveUnlinked1 = Test-Path -Path $outputlocation\onedrive-unlinked-1.txt
If ($OneDriveUnlinked1 -eq $false){
    # If OneDrive is open, close it
    $OneDriveProcess = Get-Process onedrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
    # Delete regkeys to unlink OneDrive account
    Try {
        $ErrorActionPreference = 'stop'
        Remove-Item -Path HKCU:\Software\Microsoft\OneDrive\Accounts\* -Recurse
        # If no errors deleting the registry keys, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Registry cleared: HKCU:\Software\Microsoft\OneDrive\Accounts\*"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item -Path $outputlocation\onedrive-unlinked-1.txt
		$counter ++
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-warning $errormessage
		$errormessage | out-file -append $loglocation
    }
}
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Registry already cleared: HKCU:\Software\Microsoft\OneDrive\Accounts\*"
    #write-host $output
    $output | out-file -append $loglocation
}
# Clear registry path HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\* to unlink OneDrive account 
$OneDriveUnlinked2 = Test-Path -Path $outputlocation\onedrive-unlinked-2.txt
If ($OneDriveUnlinked2 -eq $false){
    # If OneDrive is open, close it
    $OneDriveProcess = Get-Process onedrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
    # Delete regkeys to unlink OneDrive account
    Try {
        $ErrorActionPreference = 'stop'
        Remove-Item -Path HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\* -Recurse
        # If no errors deleting the registry keys, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Registry cleared: HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\*"
        #write-host $output
		$output | out-file -append $loglocation
        $null = New-Item -Path $outputlocation\onedrive-unlinked-2.txt
		$counter ++
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$errormessage=$timestamp + " ERROR: " + $_.ToString()
		#write-warning $errormessage
		$errormessage | out-file -append $loglocation
    }
}
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Registry already cleared: HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\*"
    #write-host $output
    $output | out-file -append $loglocation
}

# Begin Office Activation logout
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 4/7): Clearing Office Activation sign in"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation

$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	$OfficeCommon = Get-Item Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common -ErrorAction SilentlyContinue

	#Get the Any Identities of users.
	#Iterate through each of the Identities under Common\Identity\Identities
	#From each Identity, grab the ProviderID value, Put each into  the Identity Keys Value : SignedOutWAMUsers, seperate using semicolons
	$IdentityKey = Get-Item Registry::$OfficeCommon\Identity -ErrorAction SilentlyContinue
	$UserIdentityKeys = Get-ChildItem Registry::$IdentityKey\Identities -ErrorAction SilentlyContinue

	$UserIdentityKeysCheck = Test-Path -Path Registry::$IdentityKey\Identities
	if($UserIdentityKeysCheck -eq $False) {
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

		$UserIdentityKeys = Get-ChildItem Registry::$IdentityKey\Identities -ErrorAction SilentlyContinue
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " No users logged into Office "
		#write-host $output
		$output | out-file -append $loglocation
	}
	else {
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " UserIdentityKeys: " + $UserIdentityKeys + "\Identities"
		#write-host $output
		$output | out-file -append $loglocation
		# Iterates each Key under the Identity, and pulls the ProviderID which is an ID for each user logged in.
		$UserIdentityKeys | foreach {
			$CurrentProviderID = (Get-ItemProperty Registry::$_ -Name ProviderID -ErrorAction SilentlyContinue).ProviderID
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " CurrentProviderID: " + $CurrentProviderID
			#write-host $output
			$output | out-file -append $loglocation
			if([string]::IsNullOrWhiteSpace($SignedOutWAMUsers)) {
				$SignedOutWAMUsers = $CurrentProviderID
			}
			else {
				$SignedOutWAMUsers = $SignedOutWAMUsers +";"+ $CurrentProviderID
			}
			#Optionally, we can set a DWORD =1 for the value SignedOut on Identity SubKey
			#But they get removed next time you launch an Office program (the whole Subkey is removed)
		}
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Generated SignedOutWAMUsers: " + $SignedOutWAMUsers
		#write-host $output
		$output | out-file -append $loglocation

		#Compare the WAM to whats in the IdentityKey
		$ExistingWAM = $IdentityKey.GetValue("SignedOutWAMUsers")
		if(!($ExistingWAM -eq $null)) {
			if(!($ExistingWAM)) {
				#With the above null check, being here means empty
				$NewWam = ($SignedOutWAMUsers)
			}
			else {
				$NewWam = (($ExistingWAM +';'+ $SignedOutWAMUsers) -split ';' | Select -Unique)-join ';'
			}
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Prior WAM Existed: " + $ExistingWAM
			#write-host $output
			$output | out-file -append $loglocation
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " New WAM: " + $NewWam
			#write-host $output
			$output | out-file -append $loglocation
			Set-ItemProperty Registry::$IdentityKey -Name SignedOutWAMUsers -Value $NewWam -ErrorAction SilentlyContinue
		}
		else {
			# there wasnt one here before, so we create a new registry value and put our SignedOutWAMUsers in it.
			$null = New-ItemProperty Registry::$IdentityKey -Name SignedOutWAMUsers -Value $SignedOutWAMUsers -ErrorAction SilentlyContinue
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " No Prior WAM, Creating new Value From Generated SignOutWAMUsers"
			#write-host $output
			$output | out-file -append $loglocation		
		}
	}
}
Else {
	$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	$output = $timestamp + " SKIPPING: Office Identities already cleared"
	$output | out-file -append $loglocation	
}
    # Check if Identities Subkey was already deleted by script
    $IdentitiesKey = Test-Path -Path $outputlocation\office-identities-cleared.txt
    If($IdentitiesKey -eq $False) { 
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing Identities Registry Key: " + $IdentityKey + "\Identities"
        #write-host $output
        $output | out-file -append $loglocation
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\Identities -Recurse
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\Identities"
            #write-host $output
	        $output | out-file -append $loglocation
            $null = New-Item $outputlocation\office-identities-cleared.txt
			$counter ++
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        #write-warning $errormessage
	        $errormessage | out-file -append $loglocation
        }
    }
    # Check if Profiles Subkey was already deleted by script
    $ProfileKey = Test-Path -Path $outputlocation\office-profiles-cleared.txt
    If($ProfileKey -eq $False) { 
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing Profiles Registry Key: " + $IdentityKey + "\Profiles"
        #write-host $output
        $output | out-file -append $loglocation
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\Profiles -Recurse
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\Profiles"
            #write-host $output
	        $output | out-file -append $loglocation
            $null = New-Item $outputlocation\office-profiles-cleared.txt
			$counter ++
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        #write-warning $errormessage
	        $errormessage | out-file -append $loglocation
        }
    }
	
	#This one, I noticed starts to populate with URLS if you save to OneDrive
	#It also creates keys when Outlook the app is logged into and you select "Remember My Credentials".
    # Check if DocToIdMapping Subkey was already deleted by script
    $DocToIdMappingKey = Test-Path -Path $outputlocation\office-DocToIdMapping-cleared.txt
    If($DocToIdMappingKey -eq $False) { 
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing DocToIdMapping Registry Key: " + $IdentityKey + "\Profiles"
        #write-host $output
        $output | out-file -append $loglocation
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\DocToIdMapping\* -Recurse
             # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\DocToIdMapping\*"
            #write-host $output
	        $output | out-file -append $loglocation
            $null = New-Item $outputlocation\office-DocToIdMapping-cleared.txt
			$counter ++
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        #write-warning $errormessage
	        $errormessage | out-file -append $loglocation
        }
    }	

#Start Clearing out remanants that the user was logged in before#
$CloudPolicyKey = Test-Path -Path $outputlocation\office-cloudpolicy-cleared.txt
If($CloudPolicyKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing CloudPolicy Registry Key: " + $OfficeCommon + "\CloudPolicy"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $CloudPolicy = Get-Item Registry::$OfficeCommon\CloudPolicy
        Remove-Item Registry::$CloudPolicy\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $CloudPolicy
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-cloudpolicy-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    } 
}

$LicensingNextKey = Test-Path -Path $outputlocation\office-licensingnext-cleared.txt
If($LicensingNextKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing LicensingNext Registry Key: " + $OfficeCommon + "\Licensing\LicensingNext"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $LicensingNext = Get-Item Registry::$OfficeCommon\Licensing\LicensingNext
        Remove-Item Registry::$LicensingNext\* -Exclude "CIDtoLicenseIdsMapping" -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $LicensingNext
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-licensingnext-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    } 
}

$TemplatesKey = Test-Path -Path $outputlocation\office-templates-cleared.txt
If($TemplatesKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing Templates Registry Key: " + $OfficeCommon + "\OfficeStart\Web\Templates"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $WebTemplates = Get-Item Registry::$OfficeCommon\OfficeStart\Web\Templates
        Remove-Item Registry::$WebTemplates\* -Exclude "Anonymous" -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $WebTemplates
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-templates-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

$SettingsStoreKey = Test-Path -Path $outputlocation\office-settingsstore-cleared.txt
If($SettingsStoreKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing SettingsStore Registry Key: " + $OfficeCommon + "\Privacy\SettingsStore"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $SettingsStore =  Get-Item Registry::$OfficeCommon\Privacy\SettingsStore
        Remove-Item Registry::$SettingsStore\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $SettingsStore
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-settingsstore-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

$RoamIdKey = Test-Path -Path $outputlocation\office-roamid-cleared.txt
If($RoamIdKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing RoamId Registry Key: " + $OfficeCommon + "\Roaming\Identities"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $RoamId  = Get-Item Registry::$OfficeCommon\Roaming\Identities
        Remove-Item Registry::$RoamId\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $RoamId
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-roamid-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

$SerManCacheKey = Test-Path -Path $outputlocation\office-sermancache-cleared.txt
If($SerManCacheKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing SerManCache Registry Key: " + $OfficeCommon + "\ServicesManagerCache\Identities and \ServicesManagerCache\OnPremises"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $SerManCacheCheck = Test-Path -Path Registry::$OfficeCommon\ServicesManagerCache
        If($SerManCacheCheck -eq $true) {
            $SerManCache = Get-Item Registry::$OfficeCommon\ServicesManagerCache
            Remove-Item Registry::$SerManCache\Identities\* -Recurse
            Remove-Item Registry::$SerManCache\OnPremises\* -Recurse
            }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $SerManCache + "\Identities and \OnPremises"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-sermancache-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

#write-host "-- Clearing TargetMessaging --"
$TargetedMsgServKey = Test-Path -Path $outputlocation\office-targetedmsgserv-cleared.txt
If($TargetedMsgServKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing TargetedMsgServ Registry Key: " + $OfficeCommon + "\TargetedMessagingService\MessageData and \TargetedMessagingService\MessageMetaData"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $TargetedMsgServCheck = Test-Path -Path Registry::$OfficeCommon\TargetedMessagingService
        If($TargetedMsgServCheck -eq $true) {
            $TargetedMsgServ = Get-Item Registry::$OfficeCommon\TargetedMessagingService
            Remove-Item Registry::$TargetedMsgServ\MessageData\* -Recurse
            Remove-Item Registry::$TargetedMsgServ\MessageMetaData\* -Recurse
            }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $TargetedMsgServ + "\MessageData and \MessageMetaData"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-targetedmsgserv-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

$UrlRepkey = Test-Path -Path $outputlocation\office-urlrep-cleared.txt
If($UrlRepkey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing UrlRep Registry Key: " + $OfficeCommon + "\UrlReputation\UserPolicy"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $UrlRep = Get-Item Registry::$OfficeCommon\UrlReputation\UserPolicy
        Remove-Item Registry::$UrlRep\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $UrlRep
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-urlrep-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

#There are Registry Keys inside Office\16.0 (above Common) for each individual Office software.
#The following is used to clear user data each of these may have stored of a previous user.
#write-host "-- Clearing App Specific Keys  --"
$AppsKey = Test-Path -Path $outputlocation\office-apps-cleared.txt
If($AppsKey -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing App Specific Registry Keys"
    #write-host $output
    $output | out-file -append $loglocation
	Try {
        $ErrorActionPreference = 'stop'
        $Apps = "Access;Excel;PowerPoint;Publisher;Word;OneNote"
        $Apps -split ";" | foreach{
	        #write-host "-- -- Processing $_ -- --"
	        $CurAppKey = Get-Item Registry::"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\"$_ -ErrorAction SilentlyContinue
	        $FMRU = Get-Item Registry::$CurAppKey\'File MRU' -ErrorAction SilentlyContinue
	        $UMRU = Get-Item Registry::$CurAppKey\'User MRU' -ErrorAction SilentlyContinue
	        $PMRU = Get-Item Registry::$CurAppKey\'Place MRU' -ErrorAction SilentlyContinue
	        
	        If($FMRU -ne $NULL) {
                Remove-ItemProperty Registry::$FMRU -Name *
            }
            If($PMRU -ne$NULL) {
	            Remove-ItemProperty Registry::$PMRU -Name *
            }
	        If($UMRU -ne $NULL) {
                Remove-Item Registry::$UMRU\* -Recurse
            }

	        if($_ -eq "Word") {
		        $RLCheck = Test-Path -Path Registry::$CurAppKey\"Reading Locations"
                If($RLCheck -eq $True) {
                    Remove-Item Registry::$CurAppKey\"Reading Locations"\* -Recurse
                }
	        }
	        if($_ -eq "OneNote") {
		        $ONCheck = Test-Path -Path Registry::$CurAppKey\"OpenNotebooks"
                If($ONCheck -eq $True) {
                    Remove-Item Registry::$CurAppKey\"OpenNotebooks"\* -Recurse
                }
                $RNCheck = Test-Path -Path Registry::$CurAppKey\"RecentNotebooks"
                If($RNCheck -eq $True) {
		            Remove-Item Registry::$CurAppKey\"RecentNotebooks"\* -Recurse
                }
	        }
        }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: App Specific Registry Keys"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-apps-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}	

#Clear  Files Associated with Cached  User  data inside LocalAppData
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 5/7): Clearing LocalAppData"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation

#Most of these folders will be completely emptied, a few will have a single sub folder left behind.
$FoldersTest = Test-Path -Path $outputlocation\office-folders-cleared.txt
If($FoldersTest -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    $OfficeDataDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0"
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Clearing Office Folders: " + $OfficeDataDir
    #write-host $output
    $output | out-file -append $loglocation
    
	Try {
        $ErrorActionPreference = 'stop'
        $Folders = "aggmru;BackstageInAppNavCache;TapCache;MruServiceCache;Personalization\Content;Personalization\Governance"
        $Folders -split ";" | foreach {
            if($_ -like "Personalization\*") {
		        $SubFolderTest = Test-Path "$OfficeDataDir\$_\*"
                If($SubFolderTest -eq $True) {
                    Remove-Item "$OfficeDataDir\$_\*" -Recurse -Exclude "Anonymous"
                }
	        }
	        else {	
		        $SubFolderTest = Test-Path "$OfficeDataDir\$_\*"
                If($SubFolderTest -eq $True) {
                    Remove-Item "$OfficeDataDir\$_\*" -Recurse
                }
	        }
        }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: Office Folders " + $OfficeDataDir
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\office-folders-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

if($RmOneNoteFiles.IsPresent)
{
	$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Deleting OneNote Local Folder"
    #write-host $output
    $output | out-file -append $loglocation
    $OneNoteLocal = Test-Path -Path $outputlocation\onenote-local-cleared.txt
    If($OneNoteLocal -eq $False) { 
		# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
		$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
		If ($TeamsProcess) {
			# If 'Teams' process is running, stop it else do nothing
			$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
			Start-Sleep 3
			$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
			$output = $timestamp + " Teams process was running, so we stopped it"
			#write-host $output
			$output | out-file -append $loglocation
		}
		Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
		Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process

        Try {
            $ErrorActionPreference = 'stop'
            $ONRegKey = Get-Item Registry::"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\OneNote\Options\Save"
            $ONFolder = (Get-ItemProperty Registry::$ONRegKey).'Last Local Notebook Path'
            Remove-Item $ONFolder\* -Recurse -Force
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $ONFolder
            #write-host $output
	        $output | out-file -append $loglocation
            $null = New-Item $outputlocation\onenote-local-cleared.txt
			$counter ++
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        #write-warning $errormessage
	        $errormessage | out-file -append $loglocation
        }
    }
}

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " Clearing AppData Roaming"
#write-host $output
$output | out-file -append $loglocation

$OfficeData = Test-Path -Path $outputlocation\officedata-recent-cleared.txt
If($OfficeData -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
    Try {
        $ErrorActionPreference = 'stop'
        $OfficeDataDir = "$env:APPDATA\Microsoft\Office\Recent"
        Remove-Item $OfficeDataDir\* -Recurse 
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed recent folder: " + $OfficeDataDir
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item $outputlocation\officedata-recent-cleared.txt
		$counter ++
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}
# Clear saved credentials
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 6/7): Clearing Windows Credentials"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation

$cachedcreds =  Test-Path -Path $outputlocation\cached-creds-cleared.txt
If($cachedcreds -eq $False) { 
	# We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
	$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
	If ($TeamsProcess) {
		# If 'Teams' process is running, stop it else do nothing
		$TeamsProcess | Stop-Process -Force -ErrorAction SilentlyContinue
		Start-Sleep 3
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
		$output = $timestamp + " Teams process was running, so we stopped it"
		#write-host $output
		$output | out-file -append $loglocation
	}
	Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue | Wait-Process
	
	$Credentials = (cmdkey /list | Where-Object {$_ -like "*Target=MicrosoftOffice16_Data*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
	 }

	$Credentials = (cmdkey /list | Where-Object {$_ -like "*outlook.office365.com*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
	}

	$Credentials = (cmdkey /list | Where-Object {$_ -like "*SSO_POP_Device*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
	}
	#write-host "****** Clearing SSO_POP_Device Credentials : Complete ******"

	$Credentials = (cmdkey /list | Where-Object {$_ -like "*virtualapp/didlogical*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
	}
	#write-host "****** Clearing virtualapp/didlogical Credentials : Complete ******"

	$Credentials = (cmdkey /list | Where-Object {$_ -like "*msteams*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
	}
	#write-host "****** Clearing Teams Credentials : Complete ******"

	$Credentials = (cmdkey /list | Where-Object {$_ -like "*XboxLive*"})
	Foreach ($Target in $Credentials) {
		$Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
		$Argument = "/delete:" + $Target
		Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
		}
	$null = New-Item -Path $outputlocation\cached-creds-cleared.txt
	$counter ++
}

# Try to remove the Link School/Work account if there was one. It can be created if the first time you sign in, the user all
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING (Step 7/7): Removal of 'Link School/Work account"
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	write-host $output
}
$output | out-file -append $loglocation

# Check if School/Work account was already removed by script
$SchoolWorkAccount = Test-Path -Path $outputlocation\school-work-account-cleared.txt
If ($SchoolWorkAccount -eq $false){
    # Delete folders to remove School/Work account link
    $LocalPackagesFolder ="$env:LOCALAPPDATA\Packages"
    $AADBrokerFolder = Get-ChildItem -Path $LocalPackagesFolder -Recurse -Include "Microsoft.AAD.BrokerPlugin_*";
    $AADBrokerFolder = $AADBrokerFolder[0];
    Try {
		$ErrorActionPreference = 'stop'
        Get-ChildItem "$AADBrokerFolder\AC\TokenBroker\Accounts" | Remove-Item -Recurse -Force
        # If no errors, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Folder deleted: " + $AADBrokerFolder + "\AC\TokenBroker\Accounts"
        #write-host $output
	    $output | out-file -append $loglocation
        $null = New-Item -Path $outputlocation\school-work-account-cleared.txt
		$counter ++
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    #write-warning $errormessage
	    $errormessage | out-file -append $loglocation
    }
}

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " Script took : " + $stopwatch.Elapsed.TotalSeconds + " seconds"
$output | out-file -append $loglocation

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " Counter : " + $counter
$output | out-file -append $loglocation

# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("Migration complete, please open Outlook, Teams, and OneDrive and sign in")
}

# If most tasks complete (14/26) then mark full script as complete
# If full script complete, skip
$FullScriptCheck = Test-Path -Path $outputlocation\full-script-complete.txt
If ($FullScriptCheck -eq $false){
	If($counter -gt 14) {
		$null = New-Item -Path $outputlocation\full-script-complete.txt
	}
}



#Read-Host -Prompt "Press any key to continue"
