$p7bFile = "example.pdf.sig"
$certs = Get-PfxCertificate -FilePath $p7bFile

foreach ($cert in $certs) {
    Write-Host "Subject: $($cert.Subject)"
    Write-Host "Issuer: $($cert.Issuer)"
    Write-Host "Valid From: $($cert.NotBefore)"
    Write-Host "Valid To: $($cert.NotAfter)"
    Write-Host "--------------------------"
}