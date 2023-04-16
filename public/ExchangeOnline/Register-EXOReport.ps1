<# Mail Domain section #>
function Get-MailDomainReport {
    Write-Host "Checking domains"
    $Domains = Get-DkimSigningConfig | Select-Object -Property Id, @{Name = "Default"; Expression = { $_.IsDefault } }, @{Name = "DKIM"; Expression = { $_.Enabled } }
    if (-not ($Domains)) { $Domains = Get-AcceptedDomain | Select-Object -Property Id, "Default", @{Name = "DKIM"; Expression = { $false } } }
    $DomainsReport = @()
    foreach ($Domain in $Domains) {
        $ProcessedCount++
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Domain.Id)"
        $Domain = Get-DMARC -Domain $Domain
        $Domain = Get-SPF -Domain $Domain
        $DomainsReport += $Domain
    }
    Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Domain.Id)" -Status "Ready" -Completed
    $Report = $DomainsReport | ConvertTo-Html -As Table -Property Id, DKIM, DMARC, SPF, "DMARC record", "SPF record", "DMARC hint", "SPF hint", "Default" -Fragment -PreContent "<h3 id='EXO_DOMAIN'>Domains</h3>"
    $Report = $Report -Replace "<td>False</td><td>False</td><td>False</td>", "<td class='red'>False</td><td class='red'>False</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>False</td><td>False</td><td>True</td>", "<td class='red'>False</td><td class='red'>False</td><td>True</td>"
    $Report = $Report -Replace "<td>True</td><td>False</td><td>False</td>", "<td>True</td><td class='red'>False</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>True</td><td>False</td><td>True</td>", "<td>True</td><td class='red'>False</td><td>True</td>"
    $Report = $Report -Replace "<td>False</td><td>True</td><td>False</td>", "<td class='red'>False</td><td>True</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>False</td><td>True</td><td>True</td>", "<td class='red'>False</td><td>True</td><td>True</td>"
    $Report = $Report -Replace "<td>Should be p=reject</td>", "<td class='orange'>Should be p=reject</td>"
    $Report = $Report -Replace "<td>Not sufficiently stricth</td>", "<td class='orange'>Not sufficiently strict</td>"
    $Report = $Report -Replace "<td>Not effective enough</td>", "<td class='red'>Not effective enough</td>"
    $Report = $Report -Replace "<td>Does not protect</td>", "<td class='red'>Does not protect</td>"
    $Report = $Report -Replace "<td>No qualifier found</td>", "<td class='red'>No qualifier found</td>"
    return $Report
}
function Get-DMARC {
    param($Domain)
    $DMARCRecord = (Resolve-Dns -Query "_dmarc.$($Domain.Id)" -QueryType TXT | Select-Object -Expand Answers).Text
    if ($null -eq $DMARCRecord ) {
        $DMARC = $false
    }
    else {
        switch -Regex ($DMARCRecord ) {
            ('p=none') {
                $DmarcHint = "Does not protect"
                $DMARC = $true
            }
            ('p=quarantine') {
                $DmarcHint = "Should be p=reject"
                $DMARC = $true
            }
            ('p=reject') {
                $DmarcHint = "Will protect"
                $DMARC = $true
            }
            ('sp=none') {
                $DmarcHint += "Does not protect"
                $DMARC = $true
            }
            ('sp=quarantine') {
                $DmarcHint += "Should be p=reject"
                $DMARC = $true
            }
            ('sp=reject') {
                $DmarcHint += "Will protect"
                $DMARC = $true
            }
        }
    }
    $Domain | Add-Member NoteProperty "DMARC" $DMARC
    $Domain | Add-Member NoteProperty "DMARC record" "$($DMARCRecord )"
    $Domain | Add-Member NoteProperty "DMARC hint" $DmarcHint
    return $Domain
}
function Get-SPF {
    param($Domain)
    $SPFRecord = (Resolve-Dns -Query $Domain.Id -QueryType TXT | Select-Object -Expand Answers).Text | Where-Object { $_ -match "v=spf1" }
    if ($SPFRecord -match "redirect") {
        $redirect = $SPFRecord.Split(" ")
        $RedirectName = $redirect -match "redirect" -replace "redirect="
        $SPFRecord = (Resolve-Dns -Query $RedirectName -QueryType TXT | Select-Object -Expand Answers).Text | Where-Object { $_ -match "v=spf1" }
    }
    if ($null -eq $SPFRecord) {
        $SPF = $false
    }
    if ($SPFRecord -is [array]) {
        $SPFHint = "More than one SPF-record"
        $SPF = $true
    }
    Else {
        switch -Regex ($SPFRecord) {
            '~all' {
                $SPFHint = "Not sufficiently strict"
                $SPF = $true
            }
            '-all' {
                $SPFHint = "Sufficiently strict"
                $SPF = $true
            }
            "\?all" {
                $SPFHint = "Not effective enough"
                $SPF = $true
            }
            '\+all' {
                $SPFHint = "Not effective enough"
                $SPF = $true
            }
            Default {
                $SPFHint = "No qualifier found"
                $SPF = $true
            }
        }
    }
    $Domain | Add-Member NoteProperty "SPF" "$($SPF)"
    $Domain | Add-Member NoteProperty "SPF record" "$($SPFRecord)"
    $Domain | Add-Member NoteProperty "SPF hint" $SPFHint
    return $Domain
}

<# Mail connector section#>
function Get-MailConnectorReport {
    Write-Host "Checking mail connectors"
    if (-not ($Inbound = Get-InboundConnector)) { $InboundReport = "<br><h3 id='EXO_CONNECTOR_IN'>Inbound mail connector</h3><p>Not found</p>" }
    else { $InboundReport = $Inbound | ConvertTo-Html -As Table -Property Name, SenderDomains, SenderIPAddresses, Enabled -Fragment -PreContent "<br><h3 id='EXO_CONNECTOR_IN'>Inbound mail connector</h3>" }
    if (-not ($Outbound = Get-OutboundConnector -IncludeTestModeConnectors:$true)) { $OutboundReport = "<br><h3 id='EXO_CONNECTOR_OUT'>Outbound mail connector</h3><p>Not found</p>" }
    else { $OutboundReport = $Outbound | ConvertTo-Html -As Table -Property Name, RecipientDomains, SmartHosts, Enabled -Fragment -PreContent "<br><h3 id='EXO_CONNECTOR_OUT'>Outbound mail connector</h3>" }
    $Report = @()
    $Report += $InboundReport
    $Report += $OutboundReport
    return $Report
}

<# User mailbox section #>
function Get-UserMailboxReport {
    param(
        [System.Boolean]$Language
    )
    Write-Host "Checking user mailboxes"
    if ( -not ($Mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize:Unlimited -Properties DisplayName, UserPrincipalName)) {
        return "<br><h3 id='EXO_USER'>User mailbox</h3><p>Not found</p>"
    }
    if ($Language) {
        Update-MailboxLang -Mailbox $Mailboxes
    }
    $MailboxReport = @()
    foreach ($Mailbox in $Mailboxes) {
        $ProcessedCount++
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)"
        $MailboxReport += Get-MailboxLoginAndLocation $Mailbox
    }
    Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)" -Status "Ready" -Completed
    return $MailboxReport | ConvertTo-Html -As Table -Property UserPrincipalName, DisplayName, Language, TimeZone, LoginAllowed `
        -Fragment -PreContent "<br><h3 id='EXO_USER'>User mailbox</h3>"
}
function Update-MailboxLang {
    param(
        $Mailbox
    )
    Write-Host "Setting mailboxes language:" $script:MailboxLanguageCode "timezone:" $script:MailboxTimeZone
    $Mailbox | Set-MailboxRegionalConfiguration -LocalizeDefaultFolderName:$true -Language $script:MailboxLanguageCode -TimeZone $script:MailboxTimeZone
}

<# Shared mailbox section #>
function Get-SharedMailboxReport {
    param(
        [System.Boolean]$Language,
        [System.Boolean]$DisableLogin,
        [System.Boolean]$EnableCopy
    )
    Write-Host "Checking shared mailboxes"
    if ( -not ($Mailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited -Properties DisplayName,
            UserPrincipalName, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled)) {
        return "<br><h3 id='EXO_SHARED'>Shared mailbox</h3><p>Not found</p>"
    }
    if ($Language) { Update-MailboxLang -Mailbox $Mailboxes }
    if ($DisableLogin) { Disable-UserAccount $Mailboxes }
    if ($EnableCopy) {
        Enable-SharedMailboxEnableCopyToSent $Mailboxes
        $Mailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited -Properties DisplayName,
        UserPrincipalName, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
    }
    $MailboxReport = @()
    foreach ($Mailbox in $Mailboxes) {
        $ProcessedCount++
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)"
        $MailboxReport += Get-MailboxLoginAndLocation $Mailbox
    }
    Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)" -Status "Ready" -Completed
    $Report = $MailboxReport | ConvertTo-Html -As Table -Property UserPrincipalName, DisplayName, Language, TimeZone, MessageCopyForSentAsEnabled,
    MessageCopyForSendOnBehalfEnabled, LoginAllowed -Fragment -PreContent "<br><h3 id='EXO_SHARED'>Shared mailbox</h3>"
    $Report = $Report -Replace "<td>True</td><td>True</td><td>True</td>", "<td>True</td><td>True</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>False</td><td>False</td><td>True</td>", "<td>False</td><td>False</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>True</td><td>False</td><td>True</td>", "<td>True</td><td>False</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>False</td><td>True</td><td>True</td>", "<td>False</td><td>True</td><td class='red'>True</td>"
    return $Report
}
function Get-MailboxLoginAndLocation {
    param (
        $Mailbox
    )
    $ReginalConfig = $Mailbox | Get-MailboxRegionalConfiguration
    Add-Member -InputObject $Mailbox -NotePropertyName "Language" -NotePropertyValue $ReginalConfig.Language
    Add-Member -InputObject $Mailbox -NotePropertyName "TimeZone" -NotePropertyValue $ReginalConfig.TimeZone
    Add-Member -InputObject $Mailbox -NotePropertyName "LoginAllowed" -NotePropertyValue (Request-UserAccountStatus $Mailbox.UserPrincipalName)
    return $Mailbox
}
function Enable-SharedMailboxEnableCopyToSent {
    param(
        $Mailbox
    )
    Write-Host "Enable shared mailbox copy to sent"
    $Mailbox | Set-Mailbox -MessageCopyForSentAsEnabled $True -MessageCopyForSendOnBehalfEnabled $True
}

<# Unified mailbox section #>
function Get-UnifiedMailboxReport {
    param(
        [System.Boolean]$HideFromClient
    )
    Write-Host "Checking unified mailboxes"
    if ( -not ($Mailboxes = Get-UnifiedGroup -ResultSize Unlimited)) {
        return "<br><h3 id='EXO_UNIFIED'>Unified mailbox</h3><p>Not found</p>"
    }
    if ($HideFromClient) {
        Write-Host "Hiding unified mailboxes from outlook client"
        $Mailboxes | Set-UnifiedGroup -HiddenFromExchangeClientsEnabled:$true -HiddenFromAddressListsEnabled:$false
        $Mailboxes = Get-UnifiedGroup -ResultSize Unlimited 
    }
    return $Mailboxes | Sort-Object -Property PrimarySmtpAddress | ConvertTo-Html -As Table -Property DisplayName, PrimarySmtpAddress, HiddenFromAddressListsEnabled, HiddenFromExchangeClientsEnabled -Fragment -PreContent "<br><h3 id='EXO_UNIFIED'>Unified mailbox</h3>" -PostContent "<p>Unified groups = Microsoft 365 groups</p>"
}