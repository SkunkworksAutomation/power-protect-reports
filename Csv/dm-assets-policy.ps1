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
$ReportName = "dm-assets-policy"
$ReportOutPath = "C:\Reports\csv\output"
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).csv"

# GET ASSETS BASED ON FILTERS
$Filters = @(
    "protectionPolicyId ne null"
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
        $global:AuthObject.server

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

function get-dmassets {
<#
    .SYNOPSIS
    Get PowerProtect Data Manager assets

    .DESCRIPTION
    Get PowerProtect Data Manager assets based on filters

    .PARAMETER Filters
    An array of values used to filter the query

    .PARAMETER PageSize
    An int representing the desired number of elements per page

    .OUTPUTS
    System.Array

    .EXAMPLE
    PS> # GET ASSETS BASED ON A FILTER
    PS> $Filters = @(
    "name eq `"vc1-ubu-01`""
    )
    PS> $Assets = get-dmassets -Filters $Filters -PageSize $PageSize

    .EXAMPLE
    PS> # GET ALL ASSETS
    PS> $Assets = get-dmassets -PageSize $PageSize

    .LINK
    https://developer.dell.com/apis/4378/versions/19.16.0/reference/ppdm-public.yaml/paths/~1api~1v2~1assets/get

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
        
        $Results = @()
        
        $Endpoint = "assets"
        if($Filters.Length -gt 0) {
            $Join = ($Filters -join ' ') -replace '\s','%20' -replace '"','%22'
            $Endpoint = "$($Endpoint)?filter=$($Join)&pageSize=$($PageSize)"
        } else {
            $Endpoint = "$($Endpoint)?pageSize=$($PageSize)"
        }

        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)&queryState=BEGIN" `
        -Method GET `
        -ContentType 'application/json' `
        -Headers ($AuthObject.token) `
        -SkipCertificateCheck
        $Results += $Query.content
        
        $Page = 1
        do {
            $Token = "$($Query.page.queryState)"
            if($Page -gt 1) {
                $Token = "$($Paging.page.queryState)"
            }
            $Paging = Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)&queryState=$($Token)" `
            -Method GET `
            -ContentType 'application/json' `
            -Headers ($AuthObject.token) `
            -SkipCertificateCheck
            $Results += $Paging.content

            $Page++;
        } 
        until ($Paging.page.queryState -eq "END")

        return $Results

    } # END PROCESS
}

function get-dmprotectionpolicies {
    <#
        .SYNOPSIS
        Get PowerProtect Data Manager protection policies
        
        .DESCRIPTION
        Get PowerProtect Data Manager protection policies based on filters
    
        .PARAMETER Filters
        An array of values used to filter the query
    
        .PARAMETER PageSize
        An int representing the desired number of elements per page
    
        .OUTPUTS
        System.Array
    
        .EXAMPLE
        PS> # Get a protection policy
        PS> $Filters = @(
            "name eq `"Policy-VM01`""
        )
        PS>  $Policy = get-dmprotectionpolicies -Filters $Filters -PageSize 100
    
        .LINK
        https://developer.dell.com/apis/4378/versions/19.16.0/reference/ppdm-public.yaml/paths/~1api~1v2~1protection-policies/get
    
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
        $Endpoint = "protection-policies"

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
    function get-dmmtrees {
    <#
        .SYNOPSIS
        Get PowerProtect Data Manager protection policies
        
        .DESCRIPTION
        Get PowerProtect Data Manager protection policies based on filters
    
        .PARAMETER Filters
        An array of values used to filter the query
    
        .PARAMETER PageSize
        An int representing the desired number of elements per page
    
        .OUTPUTS
        System.Array
    
        .EXAMPLE
        PS> # Get dd mtrees
        PS>  $Mtrees = get-dmmtrees -PageSize 100
    
        .LINK
        https://developer.dell.com/apis/4378/versions/19.16.0/reference/ppdm-public.yaml/paths/~1api~1v2~1datadomain-mtrees/get
    
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
        $Endpoint = "datadomain-mtrees"

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

# ITERATE OVER THE PPDM HOSTS
$Report = @()
$Servers | ForEach-Object { 
    foreach($Retry in $Retires) {
        try {
            # CONNECT THE THE REST API
            connect-dmapi -Server $_
            Write-Host "[PowerProtect Data Manager]: Getting assets report" `
            -ForegroundColor Green
            # QUERY FOR THE ACTIVITIES
            $Query = get-dmassets -Filters $Filters -PageSize $PageSize
            
            # QUERY FOR POLICIES
            $Mtrees = get-dmmtrees -PageSize $PageSize

            # QUERY FOR POLICIES
            $Policies = get-dmprotectionpolicies -PageSize $PageSize

            foreach($Record in $Query) {
                $Policy = $Policies | Where-Object {$_.id -eq $Record.protectionPolicy.id}
                $Protection = $Policy.stages | where-object {$_.type -eq "PROTECTION"}
                $Schedule = $Protection.operations | Where-Object {$_.type -eq 'SYNTHETIC_FULL'}
                $Replication = $Policy.stages | where-object {$_.type -eq "REPLICATION"}
                $storageMtree = $Mtrees | Where-Object {$_.id -eq $Protection.target.dataTargetId}
                $replicationMtree = $Mtrees | Where-Object {$_.id -eq $Replication.target.dataTargetId}
                $assetHost = $null
                switch($Record.type) {
                    'FILE_SYSTEM' {
                        $assetHost = $Record.details.fileSystem.clusterName
                        break;
                    }
                    'KUBERNETES' {
                        $assetHost = $Record.details.k8s.inventorySourceName
                        break;
                    }
                    'NAS_SHARE' {
                        $assetHost = $Record.details.nasShare.nasServer.name
                        break;
                    }
                    'VMAX_STORAGE_GROUP' {
                        $assetHost = $Record.details.vmaxStorageGroup.coordinatingHostname
                        break;
                    }
                    'VMWARE_VIRTUAL_MACHINE'{
                        $assetHost = $Record.details.vm.hostName
                        break;
                    }
                    default {
                        $assetHost = $Record.details.database.clusterName
                        break;
                    }
                }
                
                $Object = [ordered]@{
                    assetHost = $assetHost
                    assetName = $Record.name
                    assettType = $Record.type
                    protectionPolicyName = $Record.protectionPolicy.name
                    protectionPolicyCategory = $Policy.category
                    protectionPolicySchedule = $Schedule.schedule.frequency
                    nextScheduledTime = $Record.nextScheduledTime
                    lastAvailableCopyTime = $Record.lastAvailableCopyTime
                    protectionCapacityGB = [math]::Round($Record.protectionCapacity.size /1GB,2)
                    protectionTime = $Record.protectionCapacity.time
                    retentionValue = $Protection.retention.interval
                    retentionUnit = $Protection.retention.unit
                    storageTarget = "$($storageMtree._embedded.storageSystem.name)/$($storageMtree.name)"
                    replicationTarget = "$($replicationMtree._embedded.storageSystem.name)/$($replicationMtree.name)"
                    ppdmServer = $_
                }

                $Report += New-Object -TypeName psobject -Property $Object
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