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
$ReportName = "dm-tsdm-vibs"
$ReportOutPath = "C:\Reports\csv\output"
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).csv"

# FILTER
$Filters = @(
    "viewType eq `"HOST`""
)
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
            password="$(ConvertFrom-SecureString -SecureString $Credential.password -AsPlainText)"
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
        Write-Host "$($global:AuthObject.server)" -ForegroundColor Green

    } #END PROCESS
} #END FUNCTION
    
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
} #END FUNCTION
    
function get-dmvirtualcontainers {
    <#
    .SYNOPSIS
    Get PowerProtect Data Manager virtual containers (vCenter)

    .DESCRIPTION
    Get PowerProtect Data Manager virtual containers (vCenter) based on filters

    .PARAMETER Filters
    An array of values used to filter the query

    .PARAMETER PageSize
    An int representing the desired number of elements per page

    .OUTPUTS
    System.Array

    .EXAMPLE
    PS> # Get the vCenter(s)
    PS> $Filters = @(
        "viewType eq `"HOST`""
    )
    PS>  $vCenter = get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($VMware)"}

    .EXAMPLE
    PS> # Get the datacenter
    PS> $Filters = @(
    "viewType eq `"HOST`"",
    "and parentId eq `"$($vCenter.id)`""
    )
    PS>  $Datacenter = get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($DC)"}

    .EXAMPLE
    PS> # Get a folder
    PS> $Filters = @(
    "viewType eq `"VM`"",
    "and parentId eq `"$($Datacenter.id)`""
    )
    PS>  $Folder= get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($FolderName)"}

    .EXAMPLE
    PS> # Get a cluster
    PS>  $Filters = @(
        "viewType eq `"HOST`"",
        "and parentId eq `"$($Datacenter.id)`""

    )
    $Cluster = get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($ClusterName)"}

    .EXAMPLE
    PS> # Get a resource pool
    PS> $Filters = @(
        "viewType eq `"HOST`"",
        "and parentId eq `"$($Cluster.id)`""
    )
    $Pool = get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($RP)"}
#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory=$true)]
        [array]$Filters,
        [Parameter( Mandatory=$true)]
        [int]$PageSize
    )
    begin {}
    process {
        $Page = 1
        $Results = @()
        $Endpoint = "vm-containers"

        if($Filters.Length -gt 0) {
            $Join = ($Filters -join ' ') -replace '\s','%20' -replace '"','%22'
            $Endpoint = "$($Endpoint)?filterType=vCenterInventory&filter=$($Join)&recursive=false"
        }

        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)" `
        -Method GET `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck
        $Results = $Query.content
        
        if($Query.page.totalPages -gt 1) {
            # INCREMENT THE PAGE NUMBER
            $Page++
            # PAGE THROUGH THE RESULTS
            do {
                $Paging = Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)&pageSize=$($PageSize)&page=$($Page)" `
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
    
function get-dmvibs {
<#
    .SYNOPSIS
    Get PowerProtect Data Manager virtual containers (vCenter)

    .DESCRIPTION
    Get PowerProtect Data Manager virtual containers (vCenter) based on filters

    .PARAMETER Filters
    An array of values used to filter the query

    .PARAMETER PageSize
    An int representing the desired number of elements per page

    .OUTPUTS
    System.Array

    .EXAMPLE
    PS> # Get the vCenter(s)
    PS> $Filters = @(
        "viewType eq `"HOST`""
    )
    PS>  $vCenter = get-dmvirtualcontainers -Filters $Filters -PageSize 100 | `
    where-object {$_.name -eq "$($VMware)"}

#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory=$true)]
        [String]$Id,
        [Parameter( Mandatory=$true)]
        [int]$PageSize
    )
    begin {}
    process {
        $Page = 1
        $Results = @()
        $Endpoint = "vib-details"
        $Endpoint = "$($Endpoint)?parentResourceId=$($Id)&pageSize=$($PageSize)&page=$($Page)"


        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)" `
        -Method GET `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck
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

# WORKFLOW
# ITERATE OVER THE PPDM HOSTS
$Report = @()
$Servers | ForEach-Object { 
    foreach($Retry in $Retires) {
        try {
            # CONNECT THE THE REST API
            connect-dmapi -Server $_
            Write-Host "[PowerProtect Data Manager]: Getting getting vmware installation bundle information" `
            -ForegroundColor Green
            
            # QUERY FOR THE VCENTERS
            $vCenters = get-dmvirtualcontainers -Filters $Filters -PageSize $PageSize
            
            foreach($vCenter in $vCenters) {
                # GET THE DATACENTER
                $Filters = @(
                    "viewType eq `"HOST`"",
                    "and parentId eq `"$($vCenter.id)`""
                )
                $Datacenters = get-dmvirtualcontainers -Filters $Filters -PageSize $PageSize
                foreach($Datacenter in $Datacenters) {
                    # GET A CLUSTERS
                    $Filters = @(
                        "viewType eq `"HOST`"",
                        "and parentId eq `"$($Datacenter.id)`""
            
                    )
                    $Clusters = get-dmvirtualcontainers -Filters $Filters -PageSize $PageSize
                    foreach($Cluster in $Clusters) {
                        $Vibs = get-dmvibs -Id $Cluster.id -PageSize $PageSize
                        # GET THE ESX HOSTS
                        $Filters = @(
                            "viewType eq `"HOST`"",
                            "and parentId eq `"$($Cluster.id)`""
                        )
                        $EsxHosts = get-dmvirtualcontainers -Filters $Filters -PageSize $PageSize | `
                        Where-Object {$_.type -eq "esxHost"} | Sort-Object name
                        foreach($EsxHost in $EsxHosts) {
                            $install = $vibs | Where-Object {$_.resourceId -eq $EsxHost.id}
                            $object = [ordered]@{
                                vcenter = $vCenter.name
                                datacenter = $Datacenter.name
                                cluster = $Cluster.name
                                esxhost = $EsxHost.name
                                esxhostId = $EsxHost.id
                                vibResourceId = $install.resourceId
                                vibResourceType = $install.resourceType
                                vibStatus = $install.status
                                vibVersion = $install.version
                                ppdmServer = $_
                            }
            
                            $Report += (New-Object -TypeName psobject -Property $object)
                        }
                    }
                }
            }
            
            # DISCONNECT THE THE REST API
            disconnect-dmapi
            # BREAK OUT OF THE CURRENT LOOP (RETRIES)
            break;
        } catch {
            if($Retry -lt $Retires.length) {
                Write-Host "[WARNING]: $($_). Sleeping $($Seconds) seconds... Attempt #: $($Retry)" -ForegroundColor Yellow
                Start-Sleep -Seconds $Seconds
            } else {
                Write-Host "[ERROR]: $($_). Attempts: $($Retry), moving on..." -ForegroundColor Red
            }
        }
    } # END RETRIES
}

# EXPORT TO CSV
$Report | Export-Csv $ReportOutFile -NoTypeInformation