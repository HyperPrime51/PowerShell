<#
	Connects to exchange online, takes a list of mailboxes, and exports full access, send as, send on behalf of, calendar, contact delegation and MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled settings. Exchange Online will prompt for credentials, could be modified to save credentials or use app registration to run exports
#>
# set variables
# set username
$username = "user@domain.com"

# set path for list of mailboxes
$allmbs = "c:\temp\all_mbs.csv"

# set path for export of files
$export_path_fullaccess = "c:\temp\fullaccess.csv"
$export_path_sendas = "c:\temp\sendas.csv"
$export_path_sendonbehalf = "c:\temp\sendonbehalf.csv"
$export_path_calendar = "c:\temp\calendar.csv"
$export_path_contact = "c:\temp\contact.csv"
$export_path_messagecopy = "c:\temp\messagecopy.csv"


Connect-ExchangeOnline -UserPrincipalName $username

# export current delegation for full access
$counter = 1
$array=@()
foreach ($mb in $allmbs){
	[PSCustomObject]$output=get-exomailboxpermission -identity $mb.primarysmtpaddress | Where-Object {$_.User -Notlike "*NT AUTHORITY*"} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, identity, user, @{n="Map";e={(Get-Recipient -resultsize unlimited $($_.user)).primarysmtpaddress}}, @{n="AccessRights";e={$_.accessrights -join ";"}}
	$array += $output
	Write-Progress -Activity "Exporting Full Access" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $export_path_fullaccess





# export current delegation for send as
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
foreach ($mb in $allmbs){
	[PSCustomObject]$output=get-exoRecipientPermission -identity $mb.primarysmtpaddress | Where-Object {$_.Trustee -Notlike "*NT AUTHORITY*"} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, identity, trustee, @{n="Map";e={(Get-Recipient -resultsize unlimited $($_.trustee)).primarysmtpaddress}}, @{n="AccessRights";e={$_.accessrights -join ";"}}
	$array += $output
	Write-Progress -Activity "Exporting Send As" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $export_path_sendas




# export current delegation for send on behalf of
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
foreach ($mb in $allmbs){
	$output = get-exomailbox $mb.primarysmtpaddress -properties GrantSendOnBehalfTo,emailaddresses | where {$_.GrantSendOnBehalfTo -ne $null} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, @{n="Map";e={foreach($grant in $_.GrantSendOnBehalfTo){(Get-EXORecipient $($grant)).primarysmtpaddress}}} 
	$array += $output
	Write-Progress -Activity "Exporting Send on Behalf" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | Export-CSV -notypeinfo $export_path_sendonbehalf





# export current delegation for calendars
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
ForEach($mb in $allmbs){
	[PSCustomObject]$output=Get-MailboxFolderPermission -Identity "$($mb.primarysmtpaddress):\Calendar" | Where-Object {$_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.AccessRights -ne "AvailabilityOnly"} | Select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }},User, @{n="Map";e={ $_.user.RecipientPrincipal.primarysmtpaddress}}, AccessRights
	$array += $output
	Write-Progress -Activity "Exporting Calendar" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}

$array | export-csv -notypeinfo $export_path_calendar



# export current delegation for contacts
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
ForEach($mb in $allmbs){
	[PSCustomObject]$output=Get-MailboxFolderPermission -Identity "$($mb.primarysmtpaddress):\Contacts" | Where-Object {$_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.AccessRights -ne "None"} | Select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }},User, @{n="Map";e={ $_.user.RecipientPrincipal.primarysmtpaddress}}, AccessRights
	$array += $output
	Write-Progress -Activity "Exporting Contact" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $export_path_contact



# check MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
$counter = 1
$array =@()
foreach ($mb in $allmbs){
	[PSCustomObject]$output = Get-EXOMailbox -Identity $mb.primarysmtpaddress -properties MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
	$array += $output
	Write-Progress -Activity "Exporting MessageCopy" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $export_path_messagecopy
