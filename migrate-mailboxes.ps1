<#
.SYNOPSIS
    MigrateMailboxes - Migrates linked and shared mailboxes to Exchange Online
    Created by Mike Koch on December 20, 2019
.DESCRIPTION
    Remote powershell to on-premises Exchange
        Query Exchange for linked and/or shared mailboxes to migrate
    Remote powershell to Exchange Online
        Submit migration batch request with CSV file
.NOTES
    DEPENDENCIES
    1. Functions-PSStoredCredentials.ps1 - http://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
        - contains functions to store and retrieve encrypted credentials from the local file system
        - required so that script can run unattended
    
    ASSUMPTIONS
    1. Assumes the account running this script has admin rights in the on-premises Exchange environment, as well as Account Operator
        rights in the linked domain.
    
    TO-DO
    1. Integrate my linked mailbox creation script, to result in one script that handles everything, easier to maintain
    #>

    [CmdletBinding()]
    Param()
    
    $linkedDC = "dc1.userdomain.local"  # needed only to add linked mailbox users to O365 licensing groups
    
    ## IMPORTANT: encrypted credentials can only be retrieved by the same account that was used to encrypt them
    ## AND must be on the same machine where they were encrypted
    . "C:\Scripts\Functions-PSStoredCredentials.ps1"
    $cred = Get-StoredCredential -UserName globaladmin@yourcompany.onmicrosoft.com
    
    ##### Connect to on-premises Exchange, import only the commands we plan to use
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://OnPremExchangeServer/powershell/ -Authentication Kerberos -AllowRedirection
    Import-PSSession $Session -CommandName Get-Mailbox, Set-Mailbox
    
    ##### Build a list of mailboxes to migrate
    ## Some older shared mailboxes stay on-premises, so we'll set a date variable to limit our query to recently created mailboxes
    $SharedMailboxThreshold = (Get-Date).AddDays(-30)
    ## A separate script creates linked mailboxes and sets extensionAttribute1 to "migrate.me"
    $MailboxesToMigrate = @(Get-Mailbox | Where-Object {$_.RecipientTypeDetails -like "LinkedMailbox" -AND $_.CustomAttribute1 -like "migrate.me"})
    ## Returns recently created shared mailboxes that are not already being migrated (see line 90, below)
    $MailboxesToMigrate += @(Get-Mailbox | Where-Object {$_.RecipientTypeDetails -like "SharedMailbox" -AND $_.whenMailboxCreated -gt $SharedMailboxThreshold -AND $_.CustomAttribute1 -notlike "migration in progress"})
    
    if ($MailboxesToMigrate.count -gt 0) {
        ##### Initiate remote powershell connection to Exchange Online, import only the commands needed to submit a migration batch
        $exo = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection
        Import-PSSession $exo -CommandName Get-MigrationEndPoint, New-MigrationBatch
        $MigrationEndpointOnPrem = Get-MigrationEndpoint -Identity owa.on-prem-endpoint.com
        
        foreach ($mbx in $MailboxesToMigrate) {
            ## don't try to migrate a mailbox until it's at least one hour old
            ## this ensures that Azure AD Connect has had enough time to replicate the object to Azure and Exchange Online
            if ((New-TimeSpan -Start $mbx.whenMailboxCreated -End (Get-Date).AddHours(-1)) -gt 0) {
                $mbx | Select-Object @{Name="EmailAddress";Expression={$_.primarysmtpaddress}} | Export-Csv "c:\temp\mbx.csv" -NoTypeInformation
    
                ## use the mailbox name as the migration batch name, but make sure it doesn't exceed the 64-char limit
                ## very unlikely, but costs almost nothing to do
                $batchname = "$($mbx.displayName)"
                if ($batchname.Length -gt 60) {
                    $batchname = $batchname.Substring(0,60)
                }
    
                ###### Submit the migration batch
                Write-Verbose "Submitting migration batch request..."
                New-MigrationBatch -Name $batchname `
                    -SourceEndpoint $MigrationEndpointOnPrem.Identity `
                    -TargetDeliveryDomain yourcompany.mail.onmicrosoft.com `
                    -CSVData ([System.IO.File]::ReadAllBytes("c:\temp\mbx.csv")) `
                    -NotificationEmails "EmailAdmin@yourcompany.com" `
                    -AutoStart `
                    -AutoComplete
    
                switch ($mbx.RecipientTypeDetails) {
                    "LinkedMailbox" {
                        # clear the migrate.me string from customattribute1
                        Set-Mailbox $mbx.primarySmtpAddress -CustomAttribute1 $null
                        ## we want to assign an EXO license to the user account in the linked domain (not the mailbox account)
                        ## LinkedMasterAccount contains the owner's account, in domain/username format
                        ## we just need to grab the username portion, which is the samaccountname in the linked domain
                        $sam = $mbx.LinkedMasterAccount.split("\")[1]
                        ##### Assign Office Pro Plus and Exchange Online feature licenses to the mailbox owner
                        ## Assumes use of group-based licensing, which requires an Azure AD Premium subscription
                        Add-ADGroupMember -Identity "O365 Exchange Online (E5)" -Members $sam -Server $linkedDC
                        Add-ADGroupMember -Identity "O365 Office Pro Plus (E5)" -Members $sam -Server $linkedDC
                    }
                    "SharedMailbox" {
                        ## set extensionAttribute1 to avoid adding this mailbox to future migrations (see line 44, above)
                        Set-Mailbox $mbx.primarySmtpAddress -CustomAttribute1 "migration in progress"
                    }
                    Default {}
                }
            }
        }
        Get-PSSession | Remove-PSSession
    }
    
    Write-Verbose "Finished."
    