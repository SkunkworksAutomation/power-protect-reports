# GLOBAL VARS
$global:ApiVersion = 'v2'
$global:Port = 8443
$global:AuthObject = $null

# VARS
$Servers = @(
    '10.239.100.131'
)
$Retires = @(1..3)
$Seconds = 30
$PageSize = 100

# REPORT OPTIONS
$ReportName = "dm-identities-access"
$ReportOutPath = "C:\Reports\csv\output"
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).csv"

function connect-dmapi {
<#
    .SYNOPSIS
    Connect to the PowerProtect Data Manager REST API.

    .DESCRIPTION
    Creates a credentials file for PowerProtect Data Manager if one does not exist.
    Connects to the PowerProtect Data Manager REST API

    .PARAMETER Server
    Specifies the FQDN of the PowerProtect Data Manager server.

    .OUTPUTS
    System.Object 
    $global:AuthObject

    .EXAMPLE
    PS> connect-ppdmapi -Server 'ppdm-01.vcorp.local'

    .LINK
    https://developer.dell.com/apis/4378/versions/19.16.0/docs/getting%20started/authentication-and-authorization.md

#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory=$true)]
        [string]$Server
    )
    begin {
        # CHECK TO SEE IF CREDS FILE EXISTS IF NOT CREATE ONE
        $Exists = Test-Path -Path ".\credentials\$($Server).xml" -PathType Leaf
        if($Exists) {
            $Credential = Import-CliXml ".\credentials\$($Server).xml"
        } else {
            $Credential = Get-Credential
            $Credential | Export-CliXml ".\credentials\$($Server).xml"
        } 
    }
    process {
        $Login = @{
            username="$($Credential.username)"
            password="$($Credential.GetNetworkCredential().Password)"
        }
        # LOGON TO THE POWERPROTECT API 
        $Auth = Invoke-RestMethod -Uri "https://$($Server):$($Port)/api/$($ApiVersion)/login" `
                    -Method POST `
                    -ContentType 'application/json' `
                    -Body (ConvertTo-Json $Login) `
                    -SkipCertificateCheck

        $Object = @{
            server ="https://$($Server):$($Port)/api/$($ApiVersion)"
            token= @{
                authorization="Bearer $($Auth.access_token)"
            } #END TOKEN
        } #END AUTHOBJ

        $global:AuthObject = $Object
        $global:AuthObject.server

    } #END PROCESS
} # END FUNCTION
    
function get-identityprovisions {
    <#
        .SYNOPSIS
        Get PowerProtect Data Manager identity access provisions
        
        .DESCRIPTION
        Get PowerProtect Data Manager identity access provisions
    
        .PARAMETER Filters
        An array of values used to filter the query
    
        .PARAMETER PageSize
        An int representing the desired number of elements per page
    
        .OUTPUTS
        System.Array
    
        .EXAMPLE
        PS> # Get identity-access-provisions
        PS>  $id = get-identityprovisions -PageSize 100
    
    #>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory=$false)]
        [array]$Filters,
        [Parameter( Mandatory=$true)]
        [int]$PageSize
    )
    begin {}
    process {
        
        $Page = 1
        $Results = @()
        $Endpoint = "identity-access-provisions"

        if($Filters.Length -gt 0) {
            $Join = ($Filters -join ' ') -replace '\s','%20' -replace '"','%22'
            $Endpoint = "$($Endpoint)?filter=$($Join)"
        }else {
            $Endpoint = "$($Endpoint)?"
        }
        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)pageSize=$($PageSize)&page=$($Page)" `
        -Method GET `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck

        # CAPTURE THE RESULTS
        $Results = $Query.content
        
        if($Query.page.totalPages -gt 1) {
            # INCREMENT THE PAGE NUMBER
            $Page++
            # PAGE THROUGH THE RESULTS
            do {
                $Paging = Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)pageSize=$($PageSize)&page=$($Page)" `
                -Method GET `
                -ContentType 'application/json' `
                -Headers ($AuthObject.token) `
                -SkipCertificateCheck

                # CAPTURE THE RESULTS
                $Results += $Paging.content

                # INCREMENT THE PAGE NUMBER
                $Page++   
            } 
            until ($Paging.page.number -eq $Query.page.totalPages)
        }
        return $Results

    } # END PROCESS
}

function get-identityaccessmembers {
    <#
        .SYNOPSIS
        Get PowerProtect Data Manager identity access group members
        
        .DESCRIPTION
        Get PowerProtect Data Manager identity access group members
        
        .OUTPUTS
        System.Array
    
        .EXAMPLE
        PS> # GET active-directory-identity-providers
        PS>  $members = get-identityaccessmembers

        .LINK
        https://developer.dell.com/apis/4378/versions/19.16.0/reference/ppdm-public-v2.yaml/paths/~1api~1v2~1active-directory-identity-providers~1%7Blocator%7D~1accounts~1%7Baccount-locator%7D/get
    #>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory=$false)]
        [string]$Locator,
        [Parameter( Mandatory=$false)]
        [string]$Group,
        [Parameter( Mandatory=$false)]
        [int]$PageSize
    )
    begin {}
    process {
        
        $Page = 1
        $Results = @()
        $Encode = [System.Net.WebUtility]::UrlEncode($Group)
        
        $Endpoint = "identity-sources/$($Locator)/groups/$($Group)/users"

        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)?pageSize=$($PageSize)&page=$($Page)" `
        -Method GET `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck

        # CAPTURE THE RESULTS
        $Results = $Query.content
        
        if($Query.page.totalPages -gt 1) {
            # INCREMENT THE PAGE NUMBER
            $Page++
            # PAGE THROUGH THE RESULTS
            do {
                $Paging = Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)?pageSize=$($PageSize)&page=$($Page)" `
                -Method GET `
                -ContentType 'application/json' `
                -Headers ($AuthObject.token) `
                -SkipCertificateCheck

                # CAPTURE THE RESULTS
                $Results += $Paging.content

                # INCREMENT THE PAGE NUMBER
                $Page++   
            } 
            until ($Paging.page.number -eq $Query.page.totalPages)
        }
        return $Results

    } # END PROCESS
}

function disconnect-dmapi {
<#
    .SYNOPSIS
    Disconnect from the PowerProtect Data Manager REST API.

    .DESCRIPTION
    Destroys the bearer token contained with $global:AuthObject

    .OUTPUTS
    System.Object 
    $global:AuthObject

    .EXAMPLE
    PS> disconnect-dmapi

    .LINK
    https://developer.dell.com/apis/4378/versions/19.16.0/docs/getting%20started/authentication-and-authorization.md

#>
    [CmdletBinding()]
    param (
    )
    begin {}
    process {
        #LOGOFF OF THE POWERPROTECT API
        Invoke-RestMethod -Uri "$($AuthObject.server)/logout" `
        -Method POST `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck

        $global:AuthObject = $null
    }
} # END FUNCTION

# WORKFLOW
$Report = @()
$Servers | foreach-object {
    foreach($Retry in $Retires){
        try {
            # CONNECT THE THE REST API
            connect-dmapi -Server $_
            Write-Host "[PowerProtect Data Manager]: Getting identity access report" `
            -ForegroundColor Green
            $Ids = get-identityprovisions -PageSize $PageSize

            foreach($Id in $Ids) {

                $object = [ordered]@{
                    user = "$($Id.subject)"
                    type = $Id.identityProvider.serviceMarker
                    domain = $Id.identityProvider.selector
                    roles = $Id.access.role.name -join ','
                    availableSince = $Id.availableSince
                    lastModified = $Id.lastModified
                    ppdmServer = $_
                } # END OBJECT

                $Report += (New-Object -TypeName psobject -Property $object)

                if($Id.identityProvider.serviceMarker -ne 'local'){
                    # LOOK UP THE GROUP MEMBERS
                    $members = get-identityaccessmembers -Locator $Id.identityProvider.locator -Group $Id.subject -PageSize $PageSize

                    foreach($member in $members){
                        $object = [ordered]@{
                            user = "$($Id.subject)\$($member.name)"
                            type = $Id.identityProvider.serviceMarker
                            domain = $Id.identityProvider.selector
                            roles = $Id.access.role.name -join ','
                            availableSince = $Id.availableSince
                            lastModified = $Id.lastModified
                            ppdmServer = $_
                        } # END OBJECT

                        $Report += (New-Object -TypeName psobject -Property $object)
                    }
                }
            }
            # DISCONNECT FROM THE API
            disconnect-dmapi
            break;
        }
        catch {
            if($Retry -lt $Retires.length) {
                Write-Host "[WARNING]: $($_). Sleeping $($Seconds) seconds... Attempt #: $($Retry)" -ForegroundColor Yellow
                Start-Sleep -Seconds $Seconds
            } else {
                Write-Host "[ERROR]: $($_). Attempts: $($Retry), moving on..." -ForegroundColor Red
            }
        } # END TRY / CATCH
    } # END RETRIES
} # END FOREACH

# EXPORT TO CSV
$Report | Export-Csv $ReportOutFile -NoTypeInformation