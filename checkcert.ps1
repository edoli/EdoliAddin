$CERTIFICATE_PATH = "EdoliAddIn\EdoliAddin_Temporary.pfx"
$THUMBPRINT = (Get-PfxCertificate -FilePath $CERTIFICATE_PATH).Thumbprint
Write-Host "Certificate Thumbprint: $THUMBPRINT"