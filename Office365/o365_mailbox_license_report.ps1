<#
    Connects to MSOnline and Exchange Online to pull 1 report with license and mailbox attributes. It will prompt for o365 admin creds. Could be modified to be automated by using app registration or securely saved user credentials
#>

# set account to be used to pull report
$username = "user@domain.com"

# set domain for custom column returning only email attributes of this domain
$domain = "domain.com"

# set output path
$output_path = "c:\test\mb_license_report.csv"

Import-Module MSOnline -UseWindowsPowerShell
connect-msolservice
connect-exchangeonline -userprincipalname $username

# get all mailboxes in tenant, could be modified to import a list of mailboxes from .csv
$mailbox = Get-exomailbox -resultsize unlimited -properties DisplayName, Alias, UserPrincipalName, PrimarySMTPAddress, guid, RecipientTypeDetails, WhenMailboxCreated, MessageCopyForSentAsEnabled | where {$_.userprincipalname -notlike "DiscoverySearchMailbox*"} | select DisplayName, Alias, UserPrincipalName, PrimarySMTPAddress, @{L = "$($domain)_emailaddresses"; E = { $_.emailaddresses -like "smtp:*$($domain)" -join ","}}, @{L = "onms_emailaddresses"; E = { $_.emailaddresses -like "smtp:*onmicrosoft.com" -join ","}}, guid, RecipientTypeDetails, WhenMailboxCreated, MessageCopyForSentAsEnabled

# build report
$Report=@()
$mailbox| foreach-object { 
    $DisplayName=$_.DisplayName 
    [String]$UPN=$_.UserPrincipalName
    $Alias=$_.Alias
    [String]$SmtpAddress=$_.primarysmtpaddress
    $domain_emailaddresses=$_."$($domain)_emailaddresses"
    $onms_emailaddresses=$_.onms_emailaddresses
    $GUID=$_.GUID
    $MailboxType = $_.RecipientTypeDetails
    $WhenMailboxCreated=$_.WhenMailboxCreated
    $MessageCopyForSentAsEnabled=$_.MessageCopyForSentAsEnabled
    $TotalItemSize=(get-exomailboxstatistics $SmtpAddress).TotalITemSize 
    $LastUserActionTime=(get-exomailboxstatistics $SmtpAddress).LastUserActionTime
    $LastLoggedOnUserAccount=(get-exomailboxstatistics $SmtpAddress ).LastLoggedOnUserAccount
    $msoluser=get-msoluser -userprincipalname $upn | Select userprincipalname, immutableid, BlockCredential, @{n="Licenses";e={$_.Licenses.AccountSKUid -join ";"}}
    # may need this @{n="Licenses";e={$_.LicenseAssignmentDetails.AccountSku.SkuPartNumber -join " "}}
    $Licenses=$msoluser.licenses
    $immutableID=$msoluser.immutableid
    $AccountBlocked=$msoluser.BlockCredential

    $obj=new-object System.Object
    $obj|add-member -membertype NoteProperty -name "DisplayName" -value $DisplayName
    $obj|add-member -membertype NoteProperty -name "Alias" -value $Alias
    $obj|add-member -membertype NoteProperty -name "UPN" -value $UPN
    $obj|add-member -membertype NoteProperty -name "PrimarySmtpAddress" -value $SmtpAddress
    $obj|add-member -membertype NoteProperty -name "$($domain)_Email" -value $domain_emailaddresses
    $obj|add-member -membertype NoteProperty -name "ONMS_Email" -value $onms_emailaddresses
    $obj|add-member -membertype NoteProperty -name "GUID" -value $GUID
    $obj|add-member -membertype NoteProperty -name "ImmutableID" -value $immutableID
    $obj|add-member -membertype NoteProperty -name "MailboxType" -value $MailboxType
    $obj|add-member -membertype NoteProperty -name "MessageCopyForSentAsEnabled" -value $MessageCopyForSentAsEnabled
    $obj|add-member -membertype NoteProperty -name "Licenses" -value $Licenses
    $obj|add-member -membertype NoteProperty -name "AccountBlocked" -value $AccountBlocked
    $obj|add-member -membertype NoteProperty -name "TotalItemSize" -value $TotalItemSize
    $obj|add-member -membertype NoteProperty -name "LastUserActionTime" -value $LastUserActionTime
    $obj|add-member -membertype NoteProperty -name "WhenMailboxCreated" -value $WhenMailboxCreated
    $obj|add-member -membertype NoteProperty -name "LastLogonBy" -value $LastLoggedOnUserAccount
    
    $Report+=$obj
}

$Report|export-csv $output_path

