param (
    [string]$url = ""
)

Connect-PnPOnline -Url $url
Apply-PnPProvisioningTemplate -Path .\business-docs.xml