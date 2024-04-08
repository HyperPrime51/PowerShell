<#
	This is for the source tenant
	Connects to exchange online, takes a list of mailboxes, and exports forwarding settings. Exchange Online will prompt for credentials, could be modified to save credentials or use app registration to run exports
#>

# set variables
# set username
$username = "user@domain.com"

# set path for list of mailboxes
$allmbs = "c:\temp\all_mbs.csv"

# set path for export of files
$export_path_forwarding = "c:\temp\forwarding\forwarding_export.csv"

Connect-ExchangeOnline -UserPrincipalName $username

# export email forwarding
$allmbs = import-csv $allmbs
$counter = 1
$array=@()
$output=New-Object PSObject
foreach ($forward in $allmbs){
	$output = Get-EXOMailbox $forward.primarysmtpaddress -properties primarysmtpaddress,UserPrincipalName,emailaddresses,ForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward | select primarysmtpaddress,UserPrincipalName,@{n="onms";e={$_.emailaddresses -like "smtp:*@atgf.onmicrosoft.com" -join ";" -replace "smtp:",""}},ForwardingAddress,@{n="ForwardTo";e={foreach($mb in $_.ForwardingAddress){(Get-EXORecipient -resultsize unlimited $($mb)).PrimarySmtpAddress}}},ForwardingSmtpAddress,DeliverToMailboxAndForward
	$array += $output
	Write-Progress -Activity "Exporting Forwarding" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $export_path_forwarding
