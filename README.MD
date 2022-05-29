# WebHelpDesk PowerShell Module
PowerShell module focused around interaction with SolarWinds WebHelpDesk.  Use this to manage tickets, assets & clients in an automated way.

````Powershell
Connect-WHDService -Credential (Get-Credential) -WHDURL "https://helpdesk.contoso.com"
$Asset = Get-WHDAsset -QualifierString "(networkName like `'*Server*`')" 
$Asset.assetNumber = "12345"
Update-WHDAsset -Asset $Asset
````

