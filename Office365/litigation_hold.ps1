<#  
	Connects to MSOLService, MSGraph, ExchangeOnline
	Find all mailboxes without litigation hold turned on
	Turn on litigation hold for all mailboxes that were found in previous step, collect any errors trying to turn it on
	For any mailboxes that had errors, if the error was litigation hold can't be turned on because of no license, put a license on the account, turn on litigation hold, and then take a license off.
	For any mailboxes with errors other than no license (most common, can't put a license on the mailbox because of an email address conflict), email a report
#>

$username = "enter username"
$Passwordfile = "put where password file is"
$keyfile = "put where key file is"
$key = Get-Content $KeyFile
$Credential = new-object System.Management.Automation.PSCredential $username,(Get-content $Passwordfile|ConvertTo-SecureString -key $key)

# Connect MSOLService
Import-Module MSOnline
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Connect-MSOLService -Credential $Credential

# Connect MSGraph
$AppId = "app id"
$TenantId = "tenant id"
$Certificate = Get-ChildItem Cert:\Localmachine\my\cert_generated_from_app_registration
Connect-Graph -TenantId $TenantId -AppId $AppId -Certificate $Certificate


# Connect ExchangeOnline
import-module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $credential

# Get all mailboxes that don't have lit hold turned on
$litholdneeded = Get-Mailbox "*" -Filter {LitigationHoldEnabled -eq $false} -resultsize unlimited | select guid, displayname, primarysmtpaddress, userprincipalname, @{n="emailaddresses";e={$_.emailaddresses -like "smtp:*" -join ";"}}

# Loop through and turn on lit hold. Capture errors
$need_license = @()
$other_error = @()
foreach ($i in $litholdneeded){
    try {
		Set-mailbox $i.guid -LitigationHoldEnabled $true -LitigationHoldDuration 5479 -force -ErrorAction Stop
	}
	catch {
		$error_message = $_
		$output =  new-object System.Object 
		$output | add-member -membertype NoteProperty -name "DisplayName" -value $i.DisplayName
		$output | add-member -membertype NoteProperty -name "GUID" -value $i.guid
		$output | add-member -membertype NoteProperty -name "Primarysmtpaddress" -value $i.primarysmtpaddress
		$output | add-member -membertype NoteProperty -name "Userprincipalname" -value $i.userprincipalname
		$output | add-member -membertype NoteProperty -name "EmailAddresses" -value $i.emailaddresses
		$output | add-member -membertype NoteProperty -name "Error" -value $_
		
		# Catch if error is can't turn on lit hold due to no license
		if($error_message -like "*Microsoft Exchange Online license doesn't permit you to put a litigation hold*") {
			write-host -foregroundcolor red "Mailbox $($i.primarysmtpaddress) needs license"
			$need_license += $output
		}
		# if other error message, capture
		else {
			Write-host -foregroundcolor yellow "$($i.primarysmtpaddress) other error: $_"
            $other_error += $output
		}
	}
}

# If the error was couldn't turn on lit hold due to no license, add license, turn on lit hold, then remove license
# Add license
foreach ($i in $need_license){
	try {
        Update-MgUser -UserId $i.userprincipalname -UsageLocation US -ErrorAction Stop

	}
	catch {
		write-host "Error on $($i.userprincipalname). $_"
        $output =  new-object System.Object 
		$output | add-member -membertype NoteProperty -name "DisplayName" -value $i.DisplayName
		$output | add-member -membertype NoteProperty -name "GUID" -value $i.guid
		$output | add-member -membertype NoteProperty -name "Primarysmtpaddress" -value $i.primarysmtpaddress
		$output | add-member -membertype NoteProperty -name "Userprincipalname" -value $i.userprincipalname
		$output | add-member -membertype NoteProperty -name "EmailAddresses" -value $i.emailaddresses
		$output | add-member -membertype NoteProperty -name "Error" -value $_
        $other_error += $output
	}
	try {
        Set-MgUserLicense -UserId $i.userprincipalname -AddLicenses @{SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"} -RemoveLicenses @()
	}
	catch {
		write-host "Error on $($i.userprincipalname). $_"
        $output =  new-object System.Object 
		$output | add-member -membertype NoteProperty -name "DisplayName" -value $i.DisplayName
		$output | add-member -membertype NoteProperty -name "GUID" -value $i.guid
		$output | add-member -membertype NoteProperty -name "Primarysmtpaddress" -value $i.primarysmtpaddress
		$output | add-member -membertype NoteProperty -name "Userprincipalname" -value $i.userprincipalname
		$output | add-member -membertype NoteProperty -name "EmailAddresses" -value $i.emailaddresses
		$output | add-member -membertype NoteProperty -name "Error" -value $_
        $other_error += $output
	}
}

# Wait 5 mins for O365 to register license change
Start-Sleep -seconds 300

# Loop through and turn on lit hold
foreach ($i in $need_license){
    Set-mailbox $i.guid -LitigationHoldEnabled $true -LitigationHoldDuration 5479 -force
}

# Remove license
foreach ($i in $need_license){
	# Set-MsolUserLicense -UserPrincipalName $i.userprincipalname -RemoveLicenses grate:SPE_E3
    Set-MgUserLicense -UserId $i.userprincipalname -AddLicenses @{} -RemoveLicenses @($e3Sku.skuid)
}

# If there are other errors, email report
If($other_error.count -gt 0) {
	$other_error | export-csv -notypeinfo "E:\PowerShell Scripts\Litigation Hold\output\other_error.csv"
	# email relay settings
	$smtpServer ="smtp_server"
	$SMTPPort = "25"
	$FromAddress = "From_address"
	$subject = "Litigation Hold Errors"
	$body = ""
	$textEncoding = [System.Text.Encoding]::UTF8
	$recipient = "recipient_address"

	# email report to systems
	$files = Get-ChildItem -Path "E:\PowerShell Scripts\Litigation Hold\output"
	foreach($csv in $files){
	    $attachment = "E:\PowerShell Scripts\Litigation Hold\output\$($csv.name)"
	    Send-Mailmessage -smtpServer $smtpServer -Port $SMTPPort -from $FromAddress -to $recipient -subject $subject -Attachments $attachment -Encoding $textEncoding -ErrorAction Stop
	} 
}

# Disconnect ExchangeOnline
Disconnect-ExchangeOnline -Confirm:$false
