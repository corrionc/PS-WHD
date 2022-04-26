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

function Get-WHDRequestTypes {
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
    Invoke-WHDRESTMethod -EndpointURL "AssetStatuses"
}

Function Get-WHDAssetTypes {
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
    param(
        $Asset
    )
    Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "PUT" -WHDObject $Asset
}

Function New-WHDAsset {
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
    param(
        $Asset
    )
    Invoke-WHDRESTMethod -EndpointURL "Assets/$($($Asset.id))" -Method "DELETE" -WHDObject $Asset
}
Function Get-WHDModel {
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
    param(
        $Asset
    )
    Invoke-WHDRESTMethod -EndpointURL "Models" -Method "POST" -WHDObject $Asset
}

Function Get-WHDManufacturer {
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