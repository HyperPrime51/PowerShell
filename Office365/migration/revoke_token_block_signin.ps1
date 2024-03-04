<#
    Use this script to revoke users' sign in and then disable the account.

    WARNING: This is currently set up to revoke the refresh token and disable all users in the tenant. Add admin accounts to $skip_users otherwise you can lock yourself out of the tenant.
#>

Connect-AzureAD

# *****IMPORTANT***** add admin accounts here otherwise you can lock youself out of the tenant
$skip_users = @("admin1@tenant.com","admin2@tenant.com")

# Revoke sign in
$users = Get-AzureADUser -All $true
$counter = 1
foreach ($user in $users) {
	If($skip_users.contains($user.userprincipalname) -eq $true){
		write-host "Don't revoke $($user.userprincipalname)" -ForegroundColor White
	}
	Else {
		try {
			Get-AzureADUser -objectid $user.UserPrincipalName
			# Revoke-AzureADUserAllRefreshToken -ObjectID $user.UserPrincipalName
			Write-host "Revoking token for $($user.UserPrincipalName)" -foregroundcolor green
		}
		catch {
			Write-Host "Error on $($user.UserPrincipalName): $($_)" -foregroundcolor red
		}
	    # Set-AzureADUser -ObjectID $user.userprincipalname -AccountEnabled $false
	}

	Write-Progress -Activity "Revoking token and disabling accounts" -Status "$(([math]::Round((($counter)/$users.count * 100),0))) %" -PercentComplete (($counter*100)/$users.count)
	$counter ++
}
