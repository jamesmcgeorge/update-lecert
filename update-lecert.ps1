Import-Module ACMESharp
Enable-ACMEExtensionModule -ModuleName ACMESharp.Providers.IIS -ErrorAction SilentlyContinue

$env:ACMESHARP_VAULT_PROFILE="my-vault"

$domain = "gateway.domain.com"
$iissitename = "Default Web Site"
$certname = "gateway-$(get-date -format yyyy-MM-dd--HH-mm)"

$PSEmailServer = "MAILSERVER"
$RDBroker = "rdbroker.domain.com"
$LocalEmailAddress = "email@domain.com"
$OwnerEmailAddress = "email@domain.ca"
$pfxfile = "C:\Scripts\Certs\$certname.pfx"
$CertificatePassword = "PASSWORD!"

$ErrorActionPreference = "Stop"
$EmailLog = @()

function Write-Log {
  Write-Host $args[0]
  $script:EmailLog  += $args[0]
}

Try {
    Write-Log "Generating a new identifier for $domain"
    New-ACMEIdentifier -Dns $domain -Alias $certname -VaultProfile my-vault

    Write-Log "Completing a challenge via http"
    Complete-ACMEChallenge -IdentifierRef $certname -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = $iissitename } -Force -VaultProfile my-vault

    Write-Log "Submitting the challenge"
    Submit-ACMEChallenge -IdentifierRef $certname -ChallengeType http-01 -Force -VaultProfile my-vault

    $i = 0
      do {
        $identinfo = (Update-ACMEIdentifier $certname -ChallengeType http-01 -VaultProfile my-vault).Challenges | Where-Object {$_.Status -eq "valid"}
        if($identinfo -eq $null) {
          Start-Sleep 6
          $i++
        }
      } until($identinfo -ne $null -or $i -gt 10)

      if($identinfo -eq $null) {
        Write-Log "We did not receive a completed identifier after 60 seconds"
        $Body = $EmailLog | out-string
        Send-MailMessage -SmtpServer $PSEmailServer -From $LocalEmailAddress -To $OwnerEmailAddress -Subject "Attempting to renew Let's Encrypt certificate for $domain has failed" -Body $Body
        Exit
      }

    # We now have a new identifier... so, let's create a certificate
    Write-Log "Attempting to renew Let's Encrypt certificate for $domain"

    # Generate a certificate
    Write-Log "Generating certificate for $domain"
    New-ACMECertificate $certname -Generate -Alias $certname -VaultProfile my-vault

    # Submit the certificate
    Submit-ACMECertificate $certname  -VaultProfile my-vault

    $i = 0
    do {
        $certinfo = Update-AcmeCertificate $certname -VaultProfile my-vault
        if($certinfo.SerialNumber -eq "") {
        Start-Sleep 6
        $i++
        }
    } until($certinfo.SerialNumber -ne "" -or $i -gt 10)

    if($i -gt 10) {
        Write-Log "We did not receive a completed certificate after 60 seconds"
        $Body = $EmailLog | out-string
        Send-MailMessage -SmtpServer $PSEmailServer -From $LocalEmailAddress -To $OwnerEmailAddress -Subject "Attempting to renew Let's Encrypt certificate for $domain has failed" -Body $Body
        Exit
    }

    Write-Log "Installing the new certificate to IIS"
    Install-ACMECertificate -CertificateRef $certname -Installer iis -InstallerParameters @{WebSiteRef = $iissitename; BindingHost = $domain; BindingPort = 443; CertificateFriendlyName = 'LE Gateway'; Force = $true} -VaultProfile my-vault

    Write-Log "Clearing out older Let's Encrypt certificates"
    $certRootStore = "LocalMachine"
    $certStore = "My"
    # get a date object for 5 days in the past
    $date = (Get-Date).AddDays(-5)
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($certStore,$certRootStore) 
    $store.Open('ReadWrite')
    foreach($cert in $store.Certificates) {
        if($cert.Subject -eq "CN=$domain" -And $cert.Issuer.Contains("Let's Encrypt") -And $cert.NotBefore -lt $date) {
            Write-Log "Removing certificate $($cert.Thumbprint)"
            $store.Remove($cert)
        }
    }
    $store.Close()

    # Export Certificate to PFX file
    Write-Log "Exporting the new certificate to PFX"
    Get-ACMECertificate $certname -ExportPkcs12 $pfxfile -CertificatePassword $CertificatePassword -VaultProfile my-vault
    
    # set the RD Gateway cert to the new one
    Write-Log "Updating the RD Gateway with the new certificate"
    $Password = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
    Set-RDCertificate -Role RDGateway -ImportPath $pfxfile -Password $Password -ConnectionBroker $RDBroker -Force
    Set-RDCertificate -Role RDWebAccess -ImportPath $pfxfile -Password $Password -ConnectionBroker $RDBroker -Force
    Set-RDCertificate -Role RDRedirector -ImportPath $pfxfile -Password $Password -ConnectionBroker $RDBroker -Force
    Set-RDCertificate -Role RDPublishing -ImportPath $pfxfile -Password $Password -ConnectionBroker $RDBroker -Force

    # resetting IIS
    Write-Log "Resetting IIS"
    iisreset

    # Finished
    Write-Log "Finished"
    $Body = $EmailLog | out-string
    Send-MailMessage -SmtpServer $PSEmailServer -From $LocalEmailAddress -To $OwnerEmailAddress -Subject "Let's Encrypt certificate renewed for $domain" -Body $Body

} catch {
    Write-Host $_.Exception
    $ErrorMessage = $_.Exception | format-list -force | out-string
    $EmailLog += "Let's Encrypt certificate renewal for $domain failed with exception`n$ErrorMessage`r`n`r`n"
    $Body = $EmailLog | Out-String
    Send-MailMessage -SmtpServer $PSEmailServer -From $LocalEmailAddress -To $OwnerEmailAddress -Subject "Let's Encrypt certificate renewal for $domain failed with exception" -Body $Body
    Exit
}
