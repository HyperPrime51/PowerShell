<#
    This is for the destination tenant
	Connects to exchange online, takes a list of mapping mailboxes that has source mailbox and destination mailbox, and imports forwarding settings. Exchange Online will prompt for credentials, could be modified to save credentials or use app registration to run exports
#>

# set variables
# set username
$username = "user@domain.com"

# set path for list of mapping mailboxes
$mapping = "c:\temp\mapping.csv"

# set domain variables for tenant
$onmicrosoft_domain = "*@dyoung.onmicrosoft.com"
$vanity_domain = "*@dyoung.com"

# set path for export of files from forwarding_export.ps1
$export_path_forwarding = "c:\temp\forwarding\forwarding_export.csv"

# set path for logging
$log_path = "c:\temp\forwarding\log\"

Connect-ExchangeOnline -UserPrincipalName $username

# set email forwarding
start-transcript $log_path"log_set_forwarding.txt"
$counter = 1
$forwarding = import-csv $export_path_forwarding
$mapping = import-csv $mapping
foreach($mb in $forwarding){
	# skip if no forwarding is set
	If($mb.forwardto) {
		$DeliverToMailboxAndForward = [System.Convert]::ToBoolean($mb.DeliverToMailboxAndForward)
		# find the mailbox on the source side
		$mb_found = $null
		$mb_index = 0
		foreach ($map in $mapping){
			if($map.primarysmtpaddress -eq $mb.primarysmtpaddress) {
				$mb_found = $mb_index
				break
			}
			$mb_index ++
		}
		# find the forward on the source side
		if($mb.forwardto -notlike $vanity_domain -and $mb.forwardto -notlike $onmicrosoft_domain) {
			
			$map_found = $null
			$map_index = 0
			foreach ($map in $mapping){
				if($map.primarysmtpaddress -eq $mb.forwardto) {
					$map_found = $map_index
					break
				}
				$map_index ++
			}
			if($mb_found -is [int] -and $map_found -is [int]){
				$mailbox = $mapping.dest_upn[$mb_found]
				$map = $mapping.dest_upn[$map_found]
				Try {
					$null = Set-mailbox $mailbox -DeliverToMailboxAndForward $DeliverToMailboxAndForward -ForwardingAddress $map -erroraction stop #-whatif
					write-host "SUCCESS: $($mailbox) forwarded to $($map)" -foregroundcolor green
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
					write-host -foregroundcolor red "ERROR: Map $($mb.ForwardTo) not found"
				}
			}
		}
		else {
			if($mb_found -is [int]) {
				$mailbox = $mapping.dest_upn[$mb_found]
				Try {
					$null = Set-mailbox $mailbox -DeliverToMailboxAndForward $DeliverToMailboxAndForward -ForwardingAddress $mb.forwardto -erroraction stop #-whatif
					write-host "SUCCESS: $($mailbox) forwarded to $($mb.forwardto)" -foregroundcolor green
				}
				Catch{
					Write-Host "ERROR: On $($mailbox) and $($map) $($_.ToString())" -ForegroundColor Red
				}
			}
			else {
				write-host -foregroundcolor red "ERROR: Mailbox $($mb.primarysmtpaddress) not found"
			}
		}
	}
	write-host ""
	Write-Progress -Activity "Importing Forwarding" -Status "$(([math]::Round((($counter)/$forwarding.count * 100),0))) %" -PercentComplete (($counter*100)/$forwarding.count)
	$counter ++
}
stop-transcript