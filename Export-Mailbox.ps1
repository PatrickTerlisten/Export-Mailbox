# Liste aller Mailboxen
$Mailboxes = Get-Mailbox -OrganizationalUnit "OU=lamadrama,DC=domain,DC=local" | Sort-Object -Property Alias

# Liste der fehlgeschlagenen Mailboxen
$FailedMailboxes = @()

# Begrenzt die Anzahl der zu verarbeitenden Mailboxen
# $Mailboxes = $Mailboxes[0..4]

# Hier passiert die Magie...
ForEach ($Mailbox in $Mailboxes) { 

    # Alias des Benutzers
    $MailboxAlias = ($Mailbox.Alias)

    Write-Host -ForegroundColor DarkGreen "Processing Mailbox $MailboxAlias"
    Write-Host -ForegroundColor DarkGreen `n

    # Workaround für "Verbindung zum Quellpostfach konnte nicht hergestellt werden."
    Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $false
    
    Set-CASMailbox $Mailbox -MAPIEnabled:$false
    Start-Sleep 20
    
    Set-CASMailbox $Mailbox -MAPIEnabled:$true 
    Start-Sleep 20
    
    try {
        $ErrorActionPreference = "Stop";
        
        Write-Host -ForegroundColor DarkGreen "Starting Export for Mailbox $MailboxAlias..."
        Write-Host -ForegroundColor DarkGreen `n
        
        Get-Mailbox $Mailbox | New-MailboxExportRequest -Name "Export-$MailboxAlias" -FilePath "\\server.domain.local\Exporte$\$MailboxAlias.pst" -BadItemLimit unlimited -AcceptLargeDataLoss -ErrorAction Stop
        
        Write-Host -ForegroundColor DarkGreen `n

    }

    catch {

        # Ups...
        Write-Host -ForegroundColor Red "Processing of Mailbox $MailboxAlias failed!"
        Write-Host -ForegroundColor DarkGreen `n

        # Namen merken
        $FailedMailboxes += New-Object -TypeName psobject -Property @{Mailbox="$MailboxAlias"}
        
        # aufräumen
        Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $true
        
        Set-CASMailbox $Mailbox -MAPIEnabled:$false
    
        # Nächste Mailbox
        continue

    }

    # Schleife die wartet bis der Export abgeschlossen ist
    $RequestStatus = Get-MailboxExportRequest -Name "Export-$MailboxAlias"

    while ( $RequestStatus.Status -ne "Completed" ) {

        Write-Host -ForegroundColor DarkGreen "Export for Mailbox $MailboxAlias still running..."

        $RequestStatus = Get-MailboxExportRequest -Name "Export-$MailboxAlias" ; Start-Sleep 10

    }

    # Aufräumen am Ende 
    if ( $RequestStatus.Status -eq "Completed" ) {

        Write-Host -ForegroundColor DarkGreen "Export for Mailbox $MailboxAlias done. Cleaning up..."

        Get-MailboxExportRequest -Name "Export-$MailboxAlias" | Remove-MailboxExportRequest -Confirm:$false
               
        Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $true

        Set-CASMailbox $Mailbox -MAPIEnabled:$false
    }
}

# Fehlgeschlagene Mailboxen ausgeben
Write-Host -ForegroundColor DarkGreen "Unexported mailboxes:"
Write-Host -ForegroundColor DarkGreen `n
$FailedMailboxes | Format-Table -AutoSize