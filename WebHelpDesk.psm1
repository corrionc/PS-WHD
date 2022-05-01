<#
  .SYNOPSIS
    This module contains functions for working with the SolarWinds WebHelpDesk API in PowerShell
  
  .DESCRIPTION
    
  .LINK
    API Documentation Here: http://www.solarwinds.com/documentation/webhelpdesk/docs/whd_api_12.1.0/web%20help%20desk%20api.html#common-parameters-paging
  .NOTES
    Authors: Charles Crossan, Collin Corrion, Jake Kidd
  
  .VERSION 
    1.1.0 Added Update-WHDAsset

#>

function Connect-WHDService {
    <#
    .PARAMETER username
        API UserName
    .PARAMETER Password
        API Password
    .PARAMETER WHDURL
        WebHelpDesk Base URL
#>
    param (
        [parameter(Mandatory = $true)]
        [String]
        $username,
        [String]
        $Password,
        [String]
        $apiKey,
        [Parameter(Mandatory = $true)]
        [String]
        $WHDURL
    )
    if ($apiKey) {
        $URI = "$($WHDURL)/helpdesk/WebObjects/Helpdesk.woa/ra/Session?username=$($username)&apiKey=$($apiKey)"
    }
    elseif ( $Password) {
        $URI = "$($WHDURL)/helpdesk/WebObjects/Helpdesk.woa/ra/Session?username=$($username)&password=$($Password)"
    }
    else {
        throw "APIKey or Password required"
    }

    $Response = Invoke-RestMethod -Uri $URI  -Method GET -SessionVariable session 
    Set-Variable -Scope Global -Name "WHDURL" -Value $WHDURL
    Set-Variable -Scope Global -Name "WHDSessionKey" -Value $Response.sessionKey
    Set-Variable -Scope Global -Name "WHDUsername" -Value $username
    Set-Variable -Scope Global -Name "WHDPassword" -Value $Password
    Set-Variable -Scope Global -Name "WHDapikey" -Value $apiKey
    Set-Variable -Scope Global -Name "WHDSessionKeyExpiration" -Value $(Get-Date).AddSeconds(1800)
    Set-Variable -Scope Global -Name "WHDWebSession" -Value $session
}

function Disconnect-WHDService {
    Invoke-WHDRestMethod -EndpointURL Session -Method "DELETE" 
    Clear-Variable WHDSession* -Scope Global #Clear WHDSessionKey & WHDSessionKeyExpiration
}
Function Invoke-WHDRESTMethod {
    param(
        $EndpointURL,
        $Method = "GET",
        $Page = 1,
        [System.Collections.Hashtable]
        $Parameters = @{ },
        $WHDObject,
        $Verbose
    )
    if ( test-path variable:global:"WHDURL") {
        if ( (test-path variable:global:"WHDUsername") -and ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.username = $($(Get-Variable -Name "WHDUsername").value)
        }
        elseif ( ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            throw "WHDUsername required"
        }

        if ((test-path variable:global:"WHDSessionKey") -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.sessionKey = $($(Get-Variable -Name "WHDSessionKey").value)
        }
        elseif ((test-path variable:global:"WHDapikey") -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDapikey").value))) -and -not ([string]::IsNullOrEmpty($($(Get-Variable -Name "WHDSessionKey").value)))) {
            $Parameters.apiKey = $($(Get-Variable -Name "WHDapikey").value)
        }
        elseif (test-path variable:global:"WHDPassword") {
            $Parameters.password = $($(Get-Variable -Name "WHDPassword").value)
        }
        else {
            throw "APIKey, SessionKey or Password required"
        }
    }
    else {
        throw "WHDURL Required"
    }
    $parameters.page = $Page
    $URI = "$($(Get-Variable -Name "WHDURL").Value)/helpdesk/WebObjects/Helpdesk.woa/ra/$($EndpointURL)"
    $parameterString = ($Parameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
    if ($parameterString) {
        $URI += "?$($parameterString)"
    }
    if ($Verbose) { Write-Warning  $URI }
    if (-not [string]::IsNullOrEmpty($WHDObject)) {
        $ObjectJSON = ConvertTo-Json $WHDObject -Depth 4
        Invoke-RestMethod -uri $URI -Method $Method  -WebSession $WHDWebSession -Body $ObjectJSON
    }
    else {
        Invoke-RestMethod -uri $URI -Method $Method  -WebSession $WHDWebSession
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

Function Get-WHDAssetTypes {
    <#
.SYNOPSIS
Get all types of assets. 

.DESCRIPTION
Returns every type of asset in the helpdesk (desktop, laptop, etc)

.PARAMETER AssetTypeID
Return the integer asset type (1,2 etc)

.PARAMETER QualifierString
Search using a qualifier string.  Must be escaped.

.EXAMPLE
Get-WHDAssetTypes -QualifierString "(assetType like `'*top*`')"
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

Function Get-WHDStatusTypes {
    Invoke-WHDRESTMethod -EndpointURL "StatusTypes"
}

Function Get-WHDAsset {
            <#
.SYNOPSIS
Get an asset from WebHelpDesk

.DESCRIPTION
Return every asset, a specific asset, or search based on a qualifier string 

.PARAMETER AssetID
Specific integer asset to return

.PARAMETER QualifierString
Search based on properties of the asset.  Must be escaped, returns a subset of the attributes.

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
        $limit = 10
    )
    $parameters = @{ }
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
    Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "PUT" -WHDObject $Asset
}

Function New-WHDAsset {
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
    
    Invoke-WHDRESTMethod -EndpointURL "Assets" -Method "POST" -WHDObject $Asset
}

Function Remove-WHDAsset {
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
    Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "DELETE" -WHDObject $Asset
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
    Invoke-WHDRESTMethod -EndpointURL "Models" -Method "POST" -WHDObject $Model
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
    Invoke-WHDRESTMethod -EndpointURL "Manufacturers" -Method "POST" -WHDObject $Manufacturer
}

Function Get-WHDRoom {
    param(
        $RoomID,
        $QualifierString,
        $limit = 10
    )
    $parameters = @{ }
    if ($RoomID) {
        $URI = "Rooms/$($RoomID)"
    }
    elseif ($QualifierString) {
        $parameters.qualifier = $([System.Web.HttpUtility]::UrlEncode($QualifierString))
        $URI = "Rooms"
        Write-Output "Qualifier String"
    }
    # $responses = @()
    Invoke-WHDRESTMethod -EndpointURL $URI -Parameters $parameters
}