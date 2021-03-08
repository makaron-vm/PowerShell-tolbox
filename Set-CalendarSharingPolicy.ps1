<#
.Synopsis
    Set calendar sharing permission in mailboxes where default sharing permission is different than $Sharing_Policy
    The script is provided “AS IS” with no guarantees, no warranties, and they confer no rights.

.DESCRIPTION


.NOTES
    Author: Michal Ziemba
    File Name: Set-CalendarSharingPolicy.ps1
    Version: 1.0.0, DateUpdated: 2017-03-22
    Version: 1.0.1, DateUpdated: 2017-04-11
        add try/catch for finding a $calendar and $calendar_permission



.LINK
    https://pl.linkedin.com/in/mziemba

#>
 #Credentials to connect to Exchange Online
    $Credentials = Get-AutomationPSCredential -Name 'Office 365 User Management'
    #Prefix to search for
    function Connect-ExchangeOnline {
    param (
        [System.Management.Automation.PSCredential]$Creds
    )
        #Clean up existing PowerShell Sessions
        Get-PSSession | Remove-PSSession
        #Connect to Exchange Online
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection
        Import-PSSession -Session $Session  -DisableNameChecking:$true -AllowClobber:$true  | Out-Null
    }

    Connect-ExchangeOnline -Creds $Credentials
    $VerbosePreference='Continue'


<#
 Check what  "Default Sharing Policy" is set in the Office 365 and set the $Sharing_Policy variable based on this
 The following sharing policy action values can be found:
 - CalendarSharingFreeBusySimple   Share free/busy hours only.
 - CalendarSharingFreeBusyDetail   Share free/busy hours, subject, and location.
 - CalendarSharingFreeBusyReviewer   Share free/busy hours, subject, location, and the body of the message or calendar item.
 - ContactsSharing   Share contacts only.
#>
switch (((Get-SharingPolicy).domains|Where-Object {$_ -match '\*'}) -replace "\*:","")
{

# The following roles apply specifically to calendar folders:
# AvailabilityOnly - View only availability data
# LimitedDetails   - View availability data with subject and location

    'CalendarSharingFreeBusyDetail'
        {$Sharing_Policy = "LimitedDetails"}
    Default
        {$Sharing_Policy = "AvailabilityOnly"}
}

$ErrorActionPreference = "Stop"
$mailboxes = get-mailbox -ResultSize unlimited
foreach ($mailbox in $mailboxes){
    Try
    {
        $calendar = (Get-MailboxFolderStatistics $mailbox.UserPrincipalName -FolderScope calendar | Select-Object -First 1).Identity.Replace("\",":\")
        Try
        {
            $calendar_permission = $calendar|Get-MailboxFolderPermission -user Default
        }
        catch
        {
            Write-Output "Couldn't find a default user permision for the calendar :$($calendar)`n$($_.Exception.Message)"
        }
        if ($calendar_permission.AccessRights -ne $Sharing_Policy)
        {
            try
            {
                Set-MailboxFolderPermission -Identity $calendar -User Default -AccessRights $Sharing_Policy -WarningAction stop -WhatIf
                Write-Output "Set the default calendar permission for $($mailbox.userprincipalname)"
            }
            catch
            {
                Write-Output "Failed to set the default calendar permission for $($mailbox.userprincipalname)`n$($_.Exception.Message)"
            }
        }
    }
    Catch
    {
        Write-Output "Couldn't finnd a calendar for mailbox:$($mailbox.userprincipalname)`n$($_.Exception.Message)"
    }

}

