# HelloID-Task-SA-Target-ExchangeOnline-MailboxGrantSendOnBehalf
################################################################
# Form mapping
$formObject = @{
    MailboxDistinguishedName = $form.MailboxDistinguishedName
    UsersToAdd               = $form.UsersToAdd.id
}

try {
    Write-Information "Executing ExchangeOnline action: [MailboxGrantSendOnBehalf] for: [$($formObject.MailboxDistinguishedName)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Set-Mailbox', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToAdd) {
        $null = Set-Mailbox -Identity $formObject.MailboxDistinguishedName -GrantSendOnBehalfTo @{add = "$user" } -Confirm:$false -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnline'
            TargetIdentifier  = $formObject.MailboxDistinguishedName
            TargetDisplayName = $formObject.MailboxDistinguishedName
            Message           = "ExchangeOnline action: [MailboxGrantSendOnBehalf] Added [$($user)] to mailbox [$($formObject.MailboxDistinguishedName)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnline action: [MailboxGrantSendOnBehalf] Added [$($user)] to mailbox [$($formObject.MailboxDistinguishedName)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.MailboxDistinguishedName
        TargetDisplayName = $formObject.MailboxDistinguishedName
        Message           = "Could not execute ExchangeOnline action: [MailboxGrantSendOnBehalf] for: [$($formObject.MailboxDistinguishedName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailboxGrantSendOnBehalf] for: [$($formObject.MailboxDistinguishedName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
################################################################
