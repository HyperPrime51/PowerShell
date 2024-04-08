<#
    This is for the destination tenant
	Connects to exchange online, takes a list of mapping mailboxes that has source mailbox and destination mailbox, and imports full access, send as, send on behalf of, calendar, contact delegation and MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled settings. Exchange Online will prompt for credentials, could be modified to save credentials or use app registration to run exports
#>

# set variables
# set username
$username = "user@domain.com"

# set path for list of mapping mailboxes
$mapping = "c:\temp\mapping.csv"

# set path for export of files from delegation_export.ps1
$export_path_fullaccess = "c:\temp\delegation\fullaccess_export.csv"
$export_path_sendas = "c:\temp\delegation\sendas_export.csv"
$export_path_sendonbehalf = "c:\temp\delegation\sendonbehalf_export.csv"
$export_path_calendar = "c:\temp\delegation\calendar_export.csv"
$export_path_contact = "c:\temp\delegation\contact_export.csv"
$export_path_messagecopy = "c:\temp\delegation\messagecopy_export.csv"

# set path for logging
$log_path = "c:\temp\delegation\log\"

Connect-ExchangeOnline -UserPrincipalName $username

start-transcript $log_path"log_delegation_full.txt"
$counter = 1
# set delegation for exported full access
$access_type = "FullAccess"
ForEach ($mb in $export_path_fullaccess){
	# find the mailbox on the destination side
	$mb_found = $null
	$mb_index = 0
	foreach ($map in $mapping){
		if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
			$mb_found = $mb_index
			break
		}
		$mb_index ++
	}
	# find the delegatee on the destination side
	$map_found = $null
	$map_index = 0
	foreach ($map in $mapping){
		if($mb.map.contains($map.primarysmtpaddress)) {
			$map_found = $map_index
			break
		}
		$map_index ++
	}
	if($mb_found -is [int] -and $map_found -is [int]){
		$mailbox = $mapping.dest_upn[$mb_found]
		
		$map = $mapping.dest_upn[$map_found]
		Try {
			$null = Add-MailboxPermission -Identity $mailbox -User $map -AccessRights $access_type -erroraction stop -whatif
			write-host "SUCCESS: $($map) given full access to $($mailbox)" -foregroundcolor green
		}
		Catch{
			Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
		}
	}
	else {
		if($mb_found -isnot [int]){
			write-host -foregroundcolor yello "Warning: Mailbox $($mb.primarysmtpaddress) not part of migration"
		}
		else {
			write-host -foregroundcolor red "ERROR: Map $($mb.map) not found"
		}
	}
	Write-Progress -Activity "Importing Full Access Delegation" -Status "$(([math]::Round((($counter)/$export_path_fullaccess.count * 100),0))) %" -PercentComplete (($counter*100)/$export_path_fullaccess.count)
	$counter ++
}
stop-transcript

start-transcript $log_path"log_delegation_sendas.txt"
$counter = 1
# set delegation for exported send as
$access_type = "sendas"
ForEach ($mb in $export_path_sendas){
	If($mb.map) {
		$mb_found = $null
		$mb_index = 0
		foreach ($map in $mapping){
			if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
				$mb_found = $mb_index
				break
			}
			$mb_index ++
		}
		# find the delegatee on the destination side
		$map_found = $null
		$map_index = 0
		foreach ($map in $mapping){
			if($mb.map.contains($map.primarysmtpaddress)) {
				$map_found = $map_index
				break
			}
			$map_index ++
		}
		if($mb_found -is [int] -and $map_found -is [int]){
			$mailbox = $mapping.dest_upn[$mb_found]
			
			$map = $mapping.dest_upn[$map_found]
			Try {
				$null = Add-RecipientPermission $mailbox -AccessRights $access_type -Trustee $map -Confirm:$false -erroraction stop #-whatif
				write-host "SUCCESS: $($map) given sendas to $($mailbox)" -foregroundcolor green
			}
			Catch{
				Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
			}
		}
		else {
			if($mb_found -isnot [int]){
				write-host -foregroundcolor red "ERROR: Mailbox $($mb.primarysmtpaddress) not found"
			}
			else {
				write-host -foregroundcolor red "ERROR: Map $($mb.map) not found"
			}
		}
	}
	Write-Progress -Activity "Importing Send As Delegation" -Status "$(([math]::Round((($counter)/$export_path_sendas.count * 100),0))) %" -PercentComplete (($counter*100)/$export_path_sendas.count)
	$counter ++
}
stop-transcript

start-transcript $log_path"log_delegation_sendonbehalf.txt"
$counter = 1
# set delegation for exported send as
ForEach ($mb in $export_path_sendonbehalf){
	# find the mailbox on the destination side
	$mb_found = $null
	$mb_index = 0
	foreach ($map in $mapping){
		if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
			$mb_found = $mb_index
			break
		}
		$mb_index ++
	}
	# find the delegatee on the destination side
	$map_found = $null
	$map_index = 0
	$map_sendonbehalf = @()
	foreach ($map in $mapping){
		if($mb.map.contains($map.primarysmtpaddress)) {
			$map_found = $map_index
			$map_sendonbehalf += $mapping.dest_upn[$map_found]
		}
		$map_index ++
	}
	if($mb_found -is [int] -and $map_found -is [int]){
		$mailbox = $mapping.dest_upn[$mb_found]
		Try {
			$null = Set-mailbox $mailbox -GrantSendOnBehalfTo $map_sendonbehalf -Confirm:$false -erroraction stop #-whatif
			write-host "SUCCESS: $($map_sendonbehalf) given send on behalf to $($mailbox)" -foregroundcolor green
		}
		Catch{
			Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
		}
	}
	else {
		if($mb_found -isnot [int]){
			write-host -foregroundcolor red "ERROR: Mailbox $($mb.primarysmtpaddress) not found"
		}
		else {
			write-host -foregroundcolor red "ERROR: Map $($mb.map) not found"
		}
	}
	Write-Progress -Activity "Importing Send on Behalf Delegation" -Status "$(([math]::Round((($counter)/$export_path_sendonbehalf.count * 100),0))) %" -PercentComplete (($counter*100)/$export_path_sendonbehalf.count)
	$counter ++
}
stop-transcript




start-transcript $log_path"log_delegation_calendar.txt"
$counter = 1
# apply calendar delegation
ForEach ($mb in $export_path_calendar){
	# skip if delegatee is blank
	if($mb.map) {
		# skip if mapped access is the same
		If($mb.primarysmtpaddress -ne $mb.map) {
			$accessrights=$mb.accessrights.split(" ")
			# find the mailbox on the destination side
			$mb_found = $null
			$mb_index = 0
			foreach ($map in $mapping){
				if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
					$mb_found = $mb_index
					break
				}
				$mb_index ++
			}
			# find the delegatee on the destination side
			$map_found = $null
			$map_index = 0
			foreach ($map in $mapping){
				if($mb.map.contains($map.primarysmtpaddress)) {
					$map_found = $map_index
					break
				}
				$map_index ++
			}
			if($mb_found -is [int] -and $map_found -is [int]){
				$mailbox = $mapping.dest_upn[$mb_found]
				$mailbox = "$($mailbox):\Calendar"
				$map = $mapping.dest_upn[$map_found]
				Try {
					$null = Add-MailboxFolderPermission -Identity $mailbox -User $map -AccessRights $accessrights -erroraction stop #-whatif
					Write-Host "SUCCESS: $($map) given $($accessrights) to $($mailbox)" -foregroundcolor green
				}
				Catch{
					Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
				}
			}
			else {
				if($mb_found -isnot [int]){
					write-host -foregroundcolor red "ERROR: Mailbox $($mb.primarysmtpaddress) not found"
				}
				else {
					write-host -foregroundcolor red "ERROR: Map $($mb.map) not found"
				}
			}
		}
	}
	Write-Progress -Activity "Importing Calendar Delegation" -Status "$(([math]::Round((($counter)/$export_path_calendar.count * 100),0))) %" -PercentComplete (($counter*100)/$export_path_calendar.count)
	$counter ++
}
stop-transcript


# apply contact delegation
start-transcript $log_path"log_delegation_contact.txt"
$counter = 1
ForEach ($mb in $export_path_contact){
	# skip if delegatee is blank
	if($mb.map) {
		# skip if mapped access is the same
		If($mb.primarysmtpaddress -ne $mb.map) {
			$accessrights=$mb.accessrights.split(" ")
			# find the mailbox on the destination side
			$mb_found = $null
			$mb_index = 0
			foreach ($map in $mapping){
				if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
					$mb_found = $mb_index
					break
				}
				$mb_index ++
			}
			# find the delegatee on the destination side
			$map_found = $null
			$map_index = 0
			foreach ($map in $mapping){
				if($mb.map.contains($map.primarysmtpaddress)) {
					$map_found = $map_index
					break
				}
				$map_index ++
			}
			if($mb_found -is [int] -and $map_found -is [int]){
				$mailbox = $mapping.dest_upn[$mb_found]
				$mailbox = "$($mailbox):\Contacts"
				$map = $mapping.dest_upn[$map_found]
				Try {
					$null = Add-MailboxFolderPermission -Identity $mailbox -User $map -AccessRights $accessrights -erroraction stop #-whatif
					Write-Host "SUCCESS: $($map) given $($accessrights) to $($mailbox)" -foregroundcolor green
				}
				Catch{
					Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
				}
			}
			else {
				if($mb_found -isnot [int]){
					write-host -foregroundcolor red "ERROR: Mailbox $($mb.primarysmtpaddress) not found"
				}
				else {
					write-host -foregroundcolor red "ERROR: Map $($mb.map) not found"
				}
			}
		}
	}
	Write-Progress -Activity "Importing Contact Delegation" -Status "$(([math]::Round((($counter)/$export_path_contact.count * 100),0))) %" -PercentComplete (($counter*100)/$export_path_contact.count)
	$counter ++
}
stop-transcript