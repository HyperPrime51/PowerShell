<# 
The purpose of this script is to generate autmated reports of the members of dynamic distribution groups and send them to a onedrive account for review
The steps are as follows:
    Active Directory report of all active users from 3 different Active Directory forests
    Active Directory report of all active and termed users from 3 different Active Directory forests
    Exchange online report of a preview of dynamic distribution group members using a list of groups provided at E:\PowerShell Scripts\DistributionListReport\Source\ddg.csv
    Exchange online report of a preview of static distribution group members using a list of groups provided at E:\PowerShell Scripts\DistributionListReport\Source\static.csv
    Exchange online report of AcceptMessagesOnlyFromSendersOrMembers attribute of all groups (ddg.csv and static.csv)
    Exchange online report of all "sendtoall" groups. These sendtoallgroups allow everyone in the group to send to all the distros for a certain company.
    Commented out code to email reports but left code in case functionaliy is required in the future.
    Take all reports in E:\PowerShell Scripts\DistributionListReport\Output and upload them to a onedrive account for shared viewing. 
#>

# set variables
# active directory domains to pull from
$server1 = "example1.local"
$server2 = "example2.local"
$server3 = "example3.local"

# path to export files to
$export_path = "c:\temp"

# Exchange Online connection variables
$username = "username"
$Passwordfile = "password_file"
$keyfile = "key_file"

# path to import files from
$import_path_dynamic_dg = "c:\import\ddg.csv"
$import_path_static_dg = "c:\import\static_dg.csv"
$import_path_sendtoall = "c:\import\sendtoall_groups.csv"

# Specify OneDrive Site URL and Folder name to export to
$OneDriveURL = "onedrive_url"
$DocumentLibrary = "Documents"
$TargetFolderName = 'Shared\Daily Dynamic Distribution Group Report' #Leave empty to target root folder


# export AD users and attributes
$array=@()
$output=New-Object PSObject
# get all domain 1 users
$domain1 = Get-ADUser -Filter 'enabled -eq $true' -Server $server1
foreach ($d1 in $domain1) {
	$output = Get-ADUser $d1.samaccountname -Server server -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){ 
        $array += $output
    }
}
# get all domain 2 users
$domain2 = Get-ADUser -Filter 'enabled -eq $true' -Server $server2
foreach ($d2 in $domain2) {
	$output = Get-ADUser $d2.samaccountname -Server server2 -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){
        $array += $output
    }
}

# get all domain 3 users
$domain3 = Get-ADUser -Filter 'enabled -eq $true' -Server $server3
foreach ($d3 in $domain3) {
	$output = Get-ADUser $d3.samaccountname -Server server3 -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){
        $array += $output
    }
}

# export all users to single csv
$array | export-csv -notypeinfo "$($export_path)\active_users_from_ad.csv"


# add termed users to report and export another file
$output=New-Object PSObject
# get all domain1 users
$termdomain1 = Get-ADUser -Filter 'enabled -eq $false' -Server $server1
foreach ($td1 in $termdomain1) {
	$output = Get-ADUser $td1.samaccountname -Server server -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){
        $array += $output
    }
}
# get all domain2 users
$termdomain2 = Get-ADUser -Filter 'enabled -eq $false' -Server $server2
foreach ($td2 in $termdomain2) {
	$output = Get-ADUser $td2.samaccountname -Server server2 -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){
        $array += $output
    }
}

# get all PR AD users
$termdomain3 = Get-ADUser -Filter 'enabled -eq $false' -Server $server3
foreach ($td3 in $termdomain3) {
	$output = Get-ADUser $td3.samaccountname -Server server3 -properties enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, proxyaddresses, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | select enabled, DisplayName, GivenName, MiddleName, Surname, Title, Company, EmployeeID, UserPrincipalName, mail, @{L = "proxyaddresses"; E = { $_.proxyaddresses -like "smtp:*" -join ";"}}, extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15
	If($output.surname -ne $NULL){
        $array += $output
    }
}

# export all users to single csv
$array | export-csv -notypeinfo "$($export_path)\active_and_termed_users_from_ad.csv"


# Estabilish Exchange Online connection
$key = Get-Content $KeyFile
$Credential = new-object System.Management.Automation.PSCredential $username,(Get-content $Passwordfile|ConvertTo-SecureString -key $key)
import-module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $credential

# import list of dynamic distribution groups to preview
$ddg = import-csv $import_path_dynamic_dg
# for each group, export a csv with a preview of the members
$array=@()
$output=New-Object PSObject
foreach($group in $ddg){
    $FTE = Get-DynamicDistributionGroup -Identity $group.primarysmtpaddress
	$output = Get-Recipient -resultsize unlimited -RecipientPreviewFilter ($FTE.RecipientFilter) | select @{L = "group"; E = {$group.primarysmtpaddress}}, displayname, WindowsLiveID, primarysmtpaddress, @{L = "email_aliases"; E = { $_.emailaddresses -like "smtp:*" -join ","}}, RecipientTypeDetails, custom*
	$array += $output

}
$output_file = "$($export_path)\dynamic_distribution_group_preview.csv"
$array | export-csv $output_file -notypeinfo

# import list of static distribution groups to preview
$dg = import-csv $import_path_static_dg
# for each group, export a csv with a preview of the members
$array=@()
$output=New-Object PSObject
foreach($group in $dg){
	$output = Get-DistributionGroupMember $group.primarysmtpaddress | select @{L = "group"; E = {$group.primarysmtpaddress}}, displayname, WindowsLiveID, primarysmtpaddress, @{L = "email_aliases"; E = { $_.emailaddresses -like "smtp:*" -join ","}}, RecipientTypeDetails, custom*
	$array += $output

}
$output_file = "$($export_path)\static_distribution_group_preview.csv"
$array | export-csv $output_file -notypeinfo

# for each dynamic group, export AcceptMessagesOnlyFromSendersOrMembers to hashtable
$array=@()
$output=New-Object PSObject
foreach($group in $ddg){
    $output = Get-DynamicDistributionGroup -Identity $group.primarysmtpaddress | select primarysmtpaddress, @{n="AcceptMessagesOnlyFrom";e={foreach ($user in $_.AcceptMessagesOnlyFrom){(Get-recipient $user).primarysmtpaddress }}}, @{n="AcceptMessagesOnlyFromDLMembers";e={foreach ($user in $_.AcceptMessagesOnlyFromDLMembers){(Get-recipient $user).primarysmtpaddress }}}, @{n="AcceptOnlyInternalMessages";e={$_.RequireSenderAuthenticationEnabled}} 
    $array += $output
}

# for each static group, export AcceptMessagesOnlyFromSendersOrMembers to hashtable
$output=New-Object PSObject
foreach($group in $dg){
    $output = Get-DistributionGroup -Identity $group.primarysmtpaddress | select primarysmtpaddress, @{n="AcceptMessagesOnlyFrom";e={foreach ($user in $_.AcceptMessagesOnlyFrom){(Get-recipient $user).primarysmtpaddress }}}, @{n="AcceptMessagesOnlyFromDLMembers";e={foreach ($user in $_.AcceptMessagesOnlyFromDLMembers){(Get-recipient $user).primarysmtpaddress }}}, @{n="AcceptOnlyInternalMessages";e={$_.RequireSenderAuthenticationEnabled}}
    $array += $output
}
# export AcceptMessagesOnlyFromSendersOrMembers report to csv
$output_file = "$($export_path)\dynamic_and_static_distribution_groups_acceptmessagesonlyfrom.csv"
$array | export-csv $output_file -notypeinfo

# import list of sendtoall groups to get memberships
$sendtoall = import-csv $import_path_sendtoall
$array=@()
$output=New-Object PSObject
foreach ($send in $sendtoall){
    $output = Get-DistributionGroupMember $send.primarysmtpaddress | select @{L = "group"; E = {$send.primarysmtpaddress}}, displayname, WindowsLiveID, primarysmtpaddress, @{L = "email_aliases"; E = { $_.emailaddresses -like "smtp:*" -join ","}}, RecipientTypeDetails, custom*
	$array += $output
}
# export Sendtoall groups report to csv
$output_file = "$($export_path)\send_to_all_group_membership.csv"
$array | export-csv $output_file -notypeinfo

# email relay settings
<#$smtpServer ="server"
$SMTPPort = "25"
$FromAddress = "from"
$subject = "DistributionListReport"
$body = ""
$textEncoding = [System.Text.Encoding]::UTF8
$recipient = "recipient"

# email each file to a box share
$files = Get-ChildItem -Path $export_path
foreach($csv in $files){
    $attachment = "$($export_path)\$($csv.name)"
    Send-Mailmessage -smtpServer $smtpServer -Port $SMTPPort -from $FromAddress -to $recipient -subject $subject -Attachments $attachment -Encoding $textEncoding -ErrorAction Stop
}#>

# upload files to Derek's onedrive
import-module -Name SharePointPnPPowerShellOnline

#Add required references to SharePoint client assembly to use CSOM 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")  
 
#Specify local folder path
$LocalFolder = $export_path
 
$ODCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials $username,(Get-content $Passwordfile|ConvertTo-SecureString -key $key)

$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($OneDriveURL) 
$Ctx.credentials = $ODCredential
$List = $Ctx.Web.Lists.GetByTitle("$DocumentLibrary")
$Ctx.Load($List)
$Ctx.Load($List.RootFolder)
$Ctx.ExecuteQuery()
 
#Setting Target Folder
$TargetFolder = $null;
If($TargetFolderName) {
    $TargetFolderRelativeUrl = $List.RootFolder.ServerRelativeUrl+"/"+$TargetFolderName
    $TargetFolder = $Ctx.Web.GetFolderByServerRelativeUrl($TargetFolderRelativeUrl)
    $Ctx.Load($TargetFolder)
    $Ctx.ExecuteQuery()
    if(!$TargetFolder.Exists){
    Throw  "$TargetFolderName - the target folder does not exist in the OneDrive root folder."
    }
} Else {
    $TargetFolder = $List.RootFolder
}
 
#Get files from local folder and Upload into OneDrive folder
$i = 1
$Files = (Dir $LocalFolder -File) # Read files only from root folder
#$Files = (Dir $LocalFolder -File -Recurse) # Read files both from root folder and all sub folders
$TotoalFiles = $Files.Length
ForEach ($File in $Files) {
    Try {
    Write-Progress -activity "Uplading $File" -status "$i out of $TotoalFiles completed"
    $FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $File
    $Upload = $TargetFolder.Files.Add($FileCreationInfo)
    $Ctx.Load($Upload)
    $Ctx.ExecuteQuery()
    }
    catch {
        Write-Host $_.Exception.Message -Forground "Red"
    }
$i++
}

Disconnect-ExchangeOnline -Confirm:$false