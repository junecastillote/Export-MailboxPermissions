# ==============================================================================================
# 
# NAME: Export-MailboxPermissions.ps1
# 
# AUTHOR: Tito D Castillote Jr
# DATE  : July 8, 2013
# 
# COMMENT: This script exports the permissions that exist on all mailboxes for inventory or audit purposes.
#			Result will be saved in MailboxPermissions.txt
#
# v1.0 - 
#	- Initial version
# 
# ==============================================================================================
$scriptVersion = "1.0"
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "SilentlyContinue"
Write-Host "Getting list of mailboxes. Please wait." -ForegroundColor Yellow
$finalList = @()
$i=1
$mailboxlist = Get-Mailbox -ResultSize Unlimited | ?{$_.RecipientTypeDetails -ne 'DiscoveryMailbox'} | Sort-Object Name
foreach ($mailbox in $mailboxlist) {

	Write-Progress -Activity "Reading Mailbox ($i of $($mailboxlist.count))" -status "Mailbox: $($mailbox.Name)" -percentComplete ($i / ($mailboxlist.count)*100)
	$mailPerm = Get-MailboxPermission $mailbox | where {$_.Deny -eq $false -and $_.IsInherited -eq $false}
	
			
	if (($mailPerm.count) -gt 0) {
		
		foreach ($perm in $mailPerm) {
		
			if ($perm.user.ToString() -ne 'NT AUTHORITY\SELF'){
				$tempFields = "" | Select mailboxGuid,mailboxName,mailboxDisplayName,mailboxIdentity,mailboxType,NameOfUserWithAccess,SAMOfUserWithAccess,SidOfUserWithAccess
				
				$mUser = Get-User -Identity ($perm.User.SecurityIdentifier.Value)
				$tempFields.mailboxGuid = $mailbox.ExchangeGuid
				$tempFields.mailboxName = $mailbox.Name			
				$tempFields.mailboxDisplayName = $mailbox.DisplayName
				$tempFields.mailboxIdentity = $mailbox.Identity
				$tempFields.mailboxType = $mailbox.RecipientTypeDetails
				$tempFields.NameOfUserWithAccess = $mUser.Name
				$tempFields.SAMOfUserWithAccess = $mUser.SamAccountName
				$tempFields.SidOfUserWithAccess = $perm.User.SecurityIdentifier.Value
				$finalList += $tempFields
			}			
		}		
	}
$i = $i+1
}
$finalList | Export-CSV -NoTypeInformation -Delimiter "`t" MailboxPermissions.txt
