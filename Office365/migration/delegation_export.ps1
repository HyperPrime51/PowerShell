# set variables
$onedrivepath=$env:USERPROFILE+"\OneDrive - Guaranteed Rate Inc\Documents\projects\atgf_migration\delegation\"
$onedrivepath_root=$env:USERPROFILE+"\OneDrive - Guaranteed Rate Inc\Documents\projects\atgf_migration\"




# export current delegation for full access
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array=@()
foreach ($mb in $allmbs){
	[PSCustomObject]$output=get-exomailboxpermission -identity $mb.primarysmtpaddress | Where-Object {$_.User -Notlike "*NT AUTHORITY*"} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, identity, user, @{n="Map";e={(Get-Recipient -resultsize unlimited $($_.user)).primarysmtpaddress}}, @{n="AccessRights";e={$_.accessrights -join ";"}}
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $onedrivepath"export_fullaccess_delegation.csv"





# export current delegation for send as
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
foreach ($mb in $allmbs){
	[PSCustomObject]$output=get-exoRecipientPermission -identity $mb.primarysmtpaddress | Where-Object {$_.Trustee -Notlike "*NT AUTHORITY*"} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, identity, trustee, @{n="Map";e={(Get-Recipient -resultsize unlimited $($_.trustee)).primarysmtpaddress}}, @{n="AccessRights";e={$_.accessrights -join ";"}}
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $onedrivepath"export_sendas_delegation.csv"




# export current delegation for send on behalf of
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
foreach ($mb in $allmbs){
	$output = get-exomailbox $mb.primarysmtpaddress -properties GrantSendOnBehalfTo,emailaddresses | where {$_.GrantSendOnBehalfTo -ne $null} | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, @{n="Map";e={foreach($grant in $_.GrantSendOnBehalfTo){(Get-EXORecipient $($grant)).primarysmtpaddress}}} 
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | Export-CSV -notypeinfo $onedrivepath"export_sendonbehalf_delegation.csv"





# export current delegation for calendars
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
ForEach($mb in $allmbs){
	[PSCustomObject]$output=Get-MailboxFolderPermission -Identity "$($mb.primarysmtpaddress):\Calendar" | Where-Object {$_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.AccessRights -ne "AvailabilityOnly"} | Select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }},User, @{n="Map";e={ $_.user.RecipientPrincipal.primarysmtpaddress}}, AccessRights
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}

$array | export-csv -notypeinfo $onedrivepath"export_calendar_delegation.csv"



# export current delegation for contacts
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array=@()
[PSCustomObject]$output=New-Object PSObject
ForEach($mb in $allmbs){
	[PSCustomObject]$output=Get-MailboxFolderPermission -Identity "$($mb.primarysmtpaddress):\Contacts" | Where-Object {$_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.AccessRights -ne "None"} | Select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }},User, @{n="Map";e={ $_.user.RecipientPrincipal.primarysmtpaddress}}, AccessRights
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $onedrivepath"export_contact_delegation.csv"



# check MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
$allmbs=import-csv $onedrivepath_root"atg_mapping.csv"
$counter = 1
$array =@()
foreach ($mb in $allmbs){
	[PSCustomObject]$output = Get-EXOMailbox -Identity $mb.primarysmtpaddress -properties MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled | select @{L = "primarysmtpaddress"; E = { $mb.primarysmtpaddress }}, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
	$array += $output
	Write-Progress -Activity "Working" -Status "$(([math]::Round((($counter)/$allmbs.count * 100),0))) %" -PercentComplete (($counter*100)/$allmbs.count)
	$counter ++
}
$array | export-csv -notypeinfo $onedrivepath"export_MailboxSentItemsConfiguration.csv"
