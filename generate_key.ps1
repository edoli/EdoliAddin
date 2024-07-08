$cert = New-SelfSignedCertificate `
    -Subject "CN=EdoliAddIn_Temporary" `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(5) `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyUsage DigitalSignature `
    -Type CodeSigningCert

$CertPassword = ConvertTo-SecureString -String "StringPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "EdoliAddIn/EdoliAddIn_Temporary.pfx" -Password $CertPassword