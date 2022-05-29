<#
  .SYNOPSIS
    This module contains functions for working with the SolarWinds WebHelpDesk API in PowerShell
  
  .DESCRIPTION
    
  .LINK
    API Documentation Here: https://documentation.solarwinds.com/archive/pdf/whd/whdapiguide.pdf
  .NOTES
    Authors: Charles Crossan, Collin Corrion, Jake Kidd
  
  .VERSION 
    1.1.0 Added Update-WHDAsset

#>

function Connect-WHDService {
    [CmdletBinding()]
    <#
    .PARAMETER Credential
        Credential object.  Use Get-Credential to generate
    .PARAMETER APIKey 
        API Key
    .PARAMETER WHDURL
        WebHelpDesk Base URL
    .EXAMPLE
        Connect-WHDService -Credential (Get-Credential) -WHDURL "https://helpdesk.contoso.com"
#>
    param (
        # [parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,
        [String]
        $appApiKey,
        [String]
        $apiKeyUsername,
        [Parameter(Mandatory = $true)]
        [String]
        $WHDURL
    )
    if ($Credential) {
        $URI = "$($WHDURL)/helpdesk/WebObjects/Helpdesk.woa/ra/Session?username=$($Credential.UserName)&password=$($Credential.GetNetworkCredential().Password)"
    }
    elseif ($appApiKey) {
        $URI = "$($WHDURL)/helpdesk/WebObjects/Helpdesk.woa/ra/Session?username=$($apiKeyUsername)&apiKey=$($appApiKey)"
    }
    else {
        throw "APIKey or credential required"
    }
    Write-Debug $URI

    $Response = Invoke-RestMethod -Uri $URI  -Method GET -SessionVariable session
    Set-Variable -Scope Script -Name "WHDURL" -Value $WHDURL
    Set-Variable -Scope Script -Name "WHDSessionKey" -Value $Response.sessionKey
    # Set-Variable -Scope Script -Name "WHDUsername" -Value $username
    # Set-Variable -Scope Script -Name "WHDPassword" -Value $Password
    Set-Variable -Scope Script -Name "WHDapikey" -Value $appApiKey
    Set-Variable -Scope Global -Name "WHDSessionKeyExpiration" -Value $(Get-Date).AddSeconds(1800)
    Set-Variable -Scope Script -Name "WHDWebSession" -Value $session
}

function Disconnect-WHDService {
    Invoke-WHDRestMethod -EndpointURL Session -Method "DELETE" 
    Clear-Variable WHDSession* -Scope Global #Clear WHDSessionKey & WHDSessionKeyExpiration
}
Function Invoke-WHDRESTMethod {
    [CmdletBinding()]
    param(
        $EndpointURL,
        $Method = "GET",
        $Page = 1,
        [System.Collections.Hashtable]
        $Parameters = @{ },
        $WHDObject
    )
    if ( test-path variable:script:"WHDURL") {
        if ( (test-path variable:script:"WHDUsername") -and ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.username = $($(Get-Variable -Name "WHDUsername").value)
        }
        elseif ( ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            throw "WHDUsername required"
        }

        if ((test-path variable:script:"WHDSessionKey") -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.sessionKey = $($(Get-Variable -Name "WHDSessionKey").value)
        }
        elseif ((test-path variable:script:"WHDapikey") -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDapikey").value))) -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.apiKey = $($(Get-Variable -Name "WHDapikey").value)
        }
        elseif (test-path variable:script:"WHDPassword") {
            $Parameters.password = $($(Get-Variable -Name "WHDPassword").value)
        }
        else {
            throw "APIKey, SessionKey or Credential required"
        }
    }
    else {
        throw "WHDURL Required"
    }
    $parameters.page = $Page
    $URI = "$($(Get-Variable -Name "WHDURL").Value)/helpdesk/WebObjects/Helpdesk.woa/ra/$($EndpointURL)"
    Write-Verbose "Calling $URI"
    $parameterString = ($Parameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
    if ($parameterString) {
        $URI += "?$($parameterString)"
    }
    Write-Debug "Calling $URI"
    if ($Verbose) { Write-Warning  $URI }
    if (-not [string]::IsNullOrEmpty($WHDObject)) {
        $ObjectJSON = ConvertTo-Json $WHDObject -Depth 4
        Invoke-RestMethod -uri $URI -Method $Method  -WebSession $WHDWebSession -Body $ObjectJSON
    }
    else {
        Invoke-RestMethod -uri $URI -Method $Method  -WebSession $WHDWebSession
        Set-Variable -Scope Global -Name "WHDSessionKeyExpiration" -Value $(Get-Date).AddSeconds(1800)
    }
    
    
}

function Get-WHDTicket {
    param(
        $TicketNumber,
        [ValidateSet('mine', 'group', 'flagged', 'recent')]
        $TicketList = "mine",
        $RequestTypePartialName,
        $TicketStatusType,
        $QualifierString,
        $limit = 10
    )

    $parameters = @{ }
    if ($ticketNumber) {
        $URI = "Tickets/$($ticketNumber)"
    }
    elseif ($RequestTypePartialName -or $TicketStatusType) {
       
        $QualifierStrings = @()
        $QualifierStrings += $([System.Web.HttpUtility]::UrlEncode("(problemtype.problemTypeName caseInsensitiveLike '$RequestTypePartialName')"))
        $QualifierStrings += $([System.Web.HttpUtility]::UrlEncode("(statustype.statusTypeName caseInsensitiveLike '$TicketStatusType')"))
        $parameters.qualifier = $QualifierStrings -join "and"
        $URI = "Tickets"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Tickets"
    }
    else {
        $URI = "Tickets/$($ticketList)" 
    }

    $responses = @()
    $page = 1;
    $hasMore = $true
    while ($hasMore -and $responses.count -lt $limit) {
        $temp = Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters -Page $page
        if ($temp -isnot [system.array] -or $temp.count -eq 0 ) {
            $hasMore = $false
        }
        $responses += $temp
        $page += 1 
    }


    foreach ($ticket in $responses ) {
        if ($ticket.shortDetail) {
            $ticket = Get-WHDTicket -TicketNumber $ticket.id
        }
        $ticket
    }
}

function Get-WHDRequestType {
    <#
.SYNOPSIS
Get all request types

.PARAMETER limit
Limit to first N results, 100 by default

.DESCRIPTION
Returns all service requests types.

#>

    param(
        $limit
    )
    if ($limit) {
        $parameters = @{ }
        $parameters.style = "details"
        $parameters.list = "all"
        Invoke-WHDRESTMethod -EndpointURL "RequestTypes" -Parameters $parameters
    }
    else {
        Invoke-WHDRESTMethod -EndpointURL "RequestTypes"
    }
}

Function Get-WHDClient {
        <#
.SYNOPSIS
Get a WebHelpDesk client. 

.DESCRIPTION
Returns clients matching the username

.PARAMETER UserName
An email address to search for

.EXAMPLE
Get-Client -UserName user@contoso.com
#>
    param(
        $UserName
    )
    $parameters = @{ }
    $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode( "(email caseInsensitiveLike '$UserName')"))
    
    Invoke-WHDRESTMethod -EndpointURL "Clients" -Parameters $parameters
}

Function Get-WHDAssetStatus {
    <#
.SYNOPSIS
Get Asset Statuses

.DESCRIPTION
Returns all possible asset statuses.

#>
    Invoke-WHDRESTMethod -EndpointURL "AssetStatuses"
}

Function Get-WHDAssetType {
    <#
.SYNOPSIS
Get all modifiable asset types. 

.DESCRIPTION
Returns every type of asset in the helpdesk (desktop, laptop, etc)

.PARAMETER AssetTypeID
Return the integer asset type (1,2 etc)

.PARAMETER QualifierString
Search using a qualifier string.  Must be escaped.

.EXAMPLE
Get-WHDAssetType -QualifierString "(assetType like `'*top*`')"
#>
    param(
        $AssetTypeID,
        $QualifierString
    )
    $parameters = @{ }
    if ($AssetTypeID) {
        $URI = "AssetTypes/$($AssetTypeID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "AssetTypes"
    }
    else {
        $URI = "AssetTypes"
    }
    Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters
}

Function Get-WHDStatusType {
        <#
.SYNOPSIS
Get all ticket status types. 

.DESCRIPTION
Returns every type of ticket in the helpdesk (open, closed, etc)

#>
    Invoke-WHDRESTMethod -EndpointURL "StatusTypes"
}

Function Get-WHDAsset {
    [CmdletBinding()]
    <#
.SYNOPSIS
Get an asset from WebHelpDesk

.DESCRIPTION
Return every asset, a specific asset, or search based on a qualifier string 

.PARAMETER AssetID
Specific integer asset to return

.PARAMETER QualifierString
Search based on properties of the asset.  Must be escaped, returns a subset of the attributes.

.PARAMETER Style
Specify the amount of detail to return, short or detailed

.PARAMETER Limit
Limit results to N entries, defaults to 100.

.EXAMPLE
Return every asset in WebHelpDesk ()
Get-WHDAsset

.EXAMPLE
Return a specific asset
Get-WHDAsset 2

.EXAMPLE
Return all assets with a name like 'Server'
Get-WHDAsset -QualifierString "(networkName like `'*Server*`')"
#>

    param(
        $AssetID,
        $QualifierString,
        $Style = "short",
        $limit = 10
    )
    $parameters = @{ }
    $parameters.style = $style
    if ($AssetID) {
        $URI = "Assets/$($AssetID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Assets"
    }

    $responses = @()
    $page = 1;
    $hasMore = $true
    while ($hasMore -and $responses.count -lt $limit) {
        $temp = Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters -Page $page
        if ($temp -isnot [system.array] -or $temp.count -eq 0 ) {
            $hasMore = $false
        }
        $responses += $temp
        $page += 1
    }
    #Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters
    Write-Output $responses
}
Function Update-WHDAsset {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    <#
.SYNOPSIS
Updates an existing asset

.DESCRIPTION
Updates an existing asset with new properties

.PARAMETER Asset
An Asset object, usually obtained from Get-WHDAsset

.EXAMPLE
Update-WHDAsset $UpdatedAsset
#>

    param(
        $Asset
    )
    if ($PSCmdlet.ShouldProcess("$($Asset.serialNumber)", "PUT"))
    { Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "PUT" -WHDObject $Asset }
    
}

Function New-WHDAsset {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    <#
.SYNOPSIS
Creates a new WHD asset. 

.DESCRIPTION
Creates a new WHD asset, making sure they type, manufacturer and model exist

.PARAMETER Asset
A new asset to create. 

.EXAMPLE
New-WHDAsset $NewAsset
#>

    param(
        $Asset
    )
    ##Check to make sure that AssetType,Manufacturer & Model are correct
    try { Get-WHDAssetTypes $instance.model.assetTypeId }
    catch { Write-Error "AssetType not found in WHD" -ErrorAction Stop }

    try { Get-WHDManufacturer $instance.model.manufacturerId }
    catch { Write-Error "Manufacturer not found in WHD" -ErrorAction Stop }
    
    try { Get-WHDModel $instance.model.id }
    catch { Write-Error "Model not found in WHD" -ErrorAction Stop }

    if ($PSCmdlet.ShouldProcess("$($Asset.serialNumber)", "POST")) {
        Invoke-WHDRESTMethod -EndpointURL "Assets" -Method "POST" -WHDObject $Asset
    }
}

Function Remove-WHDAsset {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    <#
.SYNOPSIS
Removes an asset from WebHelpDesk 

.DESCRIPTION
Deletes a WHD Asset (ie, sets isDeleted to True)

.PARAMETER Asset
Asset to remove from WebHelpDesk

.EXAMPLE
Remove-WHDAsset $AssetToBeDeleted
#>

    param(
        $Asset
    )
    if ($PSCmdlet.ShouldProcess("$($Asset.serialNumber)", "DELETE")) {
        Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "DELETE" -WHDObject $Asset
    }
}
Function Get-WHDModel {
    <#
.SYNOPSIS
Get a model from WebHelpDesk

.DESCRIPTION
Return every model, a specific model, or search based on a qualifier string 

.PARAMETER ModelID
Specific integer model to return

.PARAMETER QualifierString
Search based on properties of the model.  Must be escaped.

.PARAMETER Limit
Limit results to N entries

.EXAMPLE
Return every model in WebHelpDesk
Get-WHDModel

.EXAMPLE
Return a specific model
Get-WHDModel 2

.EXAMPLE
Return all models with a name like 'Mac'
Get-WHDModel -QualifierString "modelName like `'*Mac*`')"
#>
    param(
        $ModelID,
        $QualifierString,
        $limit = 10
    )
    $parameters = @{ }
    if ($ModelID) {
        $URI = "Models/$($ModelID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Models"
    }
    else {
        $URI = "Models"
    }
 
    $responses = @()
    $page = 1;
    $hasMore = $true
    while ($hasMore -and $responses.count -lt $limit) {
        $temp = Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters -Page $page
        if ($temp -isnot [system.array] -or $temp.count -eq 0 ) {
            $hasMore = $false
        }
        $responses += $temp
        $page += 1
    }
    #Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters
    Write-Output $responses
}

Function Update-WHDModel {
    Write-Output "Not implemented yet"
}

Function New-WHDModel {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    <#
.SYNOPSIS
Creates a new model in WebHelpDesk

.DESCRIPTION
Create a new model in WebHelpDesk.  The AssetType and Manufacturer must already exist. 

.PARAMETER Model
Model to create

.EXAMPLE
New-WHDModel $ModelToCreate
#>

    param(
        $Model
    )
    if ($PSCmdlet.ShouldProcess("$($Model.modelName)", "POST")) {
        Invoke-WHDRESTMethod -EndpointURL "Models" -Method "POST" -WHDObject $Model
    }
}

Function Get-WHDManufacturer {
    <#
.SYNOPSIS
Get a manufacturer from WebHelpDesk

.DESCRIPTION
Return every manufacturer, a specific one, or search based on a qualifier string 

.PARAMETER ManufacturerID
Specific integer manufacturer to return

.PARAMETER QualifierString
Search based on properties of the manufacturer.  Must be escaped.

.PARAMETER Limit
Limit results to N entries

.EXAMPLE
Return every manufacturer in WebHelpDesk
Get-WHDManufacturer

.EXAMPLE
Return a specific model
Get-WHDManufacturer

.EXAMPLE
Return all manufacturers with a name like 'Dell'
Get-WHDManufacturer -QualifierString "(name like `'*Dell*`')"
#>
    param(
        $ManufacturerID,
        $QualifierString,
        $limit = 10
    )
    $parameters = @{ }
    if ($ManufacturerID) {
        $URI = "Manufacturers/$($ManufacturerID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Manufacturers"
    }
    else {
        $URI = "Manufacturers"
    }
 
    $responses = @()
    $page = 1;
    $hasMore = $true
    while ($hasMore -and $responses.count -lt $limit) {
        $temp = Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters -Page $page
        if ($temp -isnot [system.array] -or $temp.count -eq 0 ) {
            $hasMore = $false
        }
        $responses += $temp
        $page += 1
    }
    Write-Output $responses
}
Function New-WHDManufacturer {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    <#
.SYNOPSIS
Creates a new manufacturer in WebHelpDesk

.DESCRIPTION
Create a new manufacturer in WebHelpDesk.

.PARAMETER Manufacturer
Manufacturer to create

.EXAMPLE
New-WHDManufacturer $ManufacturerToCreate
#>

    param(
        $Manufacturer
    )
    if ($PSCmdlet.ShouldProcess("$($Manufacturer.fullName)", "POST"))
    { Invoke-WHDRESTMethod -EndpointURL "Manufacturers" -Method "POST" -WHDObject $Manufacturer }

}

Function Get-WHDRoom {
    [CmdletBinding()]
    <#
.SYNOPSIS
Get information about the rooms (where it happened)

.PARAMETER RoomID
Return a specific room

.PARAMETER QualifierString
Search based on properties of the room.  Must be escaped.

.PARAMETER Limit
Limit results to N entries

.DESCRIPTION
Returns rooms defined in WebHelpDesk.

#>

    param(
        $RoomID,
        $QualifierString,
        $limit = 100
    )
    $parameters = @{ }
    if ($RoomID) {
        $URI = "Rooms/$($RoomID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Rooms"
    }
    else {
        $URI = "Rooms"
    }
    $parameters.limit = $limit
    $responses = @()
    $page = 1;
    $hasMore = $true
    while ($hasMore -and $responses.count -lt $limit) {
        $temp = Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters -Page $page
        if ($temp -isnot [system.array] -or $temp.count -eq 0 ) {
            $hasMore = $false
        }
        $responses += $temp
        $page += 1
    }

    
}