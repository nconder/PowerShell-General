# Microsoft Office Online Server script
# ote that OOS does not have the ability to set certificate by thumbprint.
# This script enables the use of the newly retrieved and stored certificate with common Office Online Server services
# For more script info see https://docs.certifytheweb.com/docs/script-hooks.html

Import-Module -Name OfficeWebApps 

$OOSUrl = "ENTER OOS URL"

# Get new cert FriendlyName to set cert in OOS
$certFN = dir cert: -Recurse | Where-Object { $_.Thumbprint -eq "$result.ManagedItem.CertificateThumbprintHash" } | Select-Object FriendlyName

# Set new cert
new-OfficeWebAppsFarm -InternalUrl $OOSUrl -ExternalUrl $OOSUrl -CertificateName $results -EditingEnabled -OfficeAddinEnabled -OnlinePictureEnabled -OnlineVideoEnabled -ClipartEnabled -TranslationEnabled -AllowOutboundHttp -ExcelAbortOnRefreshOnOpenFail $true -ExcelAllowExternalData
