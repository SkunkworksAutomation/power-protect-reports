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
$ReportName = "dm-activities-summary"
$ReportOutPath = "C:\Reports\csv\output"
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).csv"

# GET ACTIVITIES BASED ON FILTERS
$Date = (Get-Date).AddDays(-1)
$Filters = @(
    "classType eq `"JOB`"",
    "and category eq `"PROTECT`"",
    "and startTime ge `"$($Date.ToString('yyyy-MM-ddThh:mm:ss.fffZ'))`""
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
    
    function get-dmactivities {
    <#
        .SYNOPSIS
        Get PowerProtect Data Manager activities
    
        .DESCRIPTION
        Get PowerProtect Data Manager activities based on filters
    
        .PARAMETER Filters
        An array of values used to filter the query
    
        .PARAMETER PageSize
        An int representing the desired number of elements per page
    
        .OUTPUTS
        System.Array
    
        .EXAMPLE
        PS> # GET ACTIVITIES BASED ON A FILTER
        PS> $Date = (Get-Date).AddDays(-1)
        PS> $Filters = @(
        "classType eq `"JOB`""
        "and category eq `"PROTECT`""
        "and startTime ge `"$($Date.ToString('yyyy-MM-dd'))T00:00:00.000Z`""
        "and result.status eq `"FAILED`""
        )
        PS> $Activities = get-dmactivities -Filters $Filters -PageSize $PageSize
    
        .EXAMPLE
        PS> # GET ALL ACTIVITIES
        PS> $Activities = get-dmactivities -PageSize $PageSize
    
        .LINK
        https://developer.dell.com/apis/4378/versions/19.16.0/reference/ppdm-public.yaml/paths/~1api~1v2~1activities/get
    
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
            $Endpoint = "activities"
            
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
        }
    }
    
    function Convert-BytesToSize
    {
    <#
        .SYNOPSIS
        Converts any integer size given to a user friendly size.
        
        .DESCRIPTION
        Converts any integer size given to a user friendly size.
    
        .PARAMETER size
        Used to convert into a more readable format.
        Required Parameter
    
        .EXAMPLE
        Convert-BytesToSize -Size 134217728
        Converts size to show 128MB
    
        .LINK
        https://learn-powershell.net/2010/08/29/convert-bytes-to-highest-available-unit/
    #>
    
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$false,Position=0)]
        [int64]$Size
    )
    
    # DETERMINE SIZE IN BASE2
    switch ($Size)
    {
        {$Size -gt 1PB}
        {
            $NewSize = @{size=$([math]::Round(($Size /1PB),1));uom="PB"}
            Break;
        }
        {$Size -gt 1TB}
        {
            $NewSize = @{size=$([math]::Round(($Size /1TB),1));uom="TB"}
            Break;
        }
        {$Size -gt 1GB}
        {
            $NewSize = @{size=$([math]::Round(($Size /1GB),1));uom="GB"}
            Break;
        }
        {$Size -gt 1MB}
        {
            $NewSize = @{size=$([math]::Round(($Size /1MB),1));uom="MB"}
            Break;
        }
        {$Size -gt 1KB}
        {
            $NewSize = @{size=$([math]::Round(($Size /1KB),1));uom="KB"}
            Break;
        }
        Default
        {
            $NewSize = @{size=$([math]::Round($Size,2));uom="Bytes"}
            Break;
        }
    }
        return $NewSize
    
    }
    
    # ITERATE OVER THE PPDM HOSTS
    $Activities = @()
    $Servers | ForEach-Object { 
        foreach($Retry in $Retires) {
            try {
                # CONNECT THE THE REST API
                connect-dmapi -Server $_
                Write-Host "[PowerProtect Data Manager]: Getting activity summary report with a startTime > $($Date.ToString('yyyy-MM-ddThh:mm:ss.fffZ'))" `
                -ForegroundColor Green
                # QUERY FOR THE ACTIVITIES
                $Query = get-dmactivities -Filters $Filters -PageSize $PageSize
                
                # GET A UNIQUE LIST OF ASSET ID FROM THE $Query VAR
                $Unique = $Query | Select-Object @{n="assetId";e={$_.asset.id}} -Unique
    
                foreach($Item in $Unique) {
                    # GET ALL OF THE RECORDS FOR THE CURRENT ASSET ID
                    $Records = $Query | Where-Object {$Item.assetId -eq $_.asset.id}
    
                    $asset = Convert-BytesToSize -Size $Records[-1].stats.assetSizeInBytes
    
                    # GET CATEGORY COUNTS FOR EACH result.status in $Records
                    $okBackups = $Records | where-object {$_.result.status -eq "OK"}
                    $okWithErrors = $Records | where-object {$_.result.status -eq "OK_WITH_ERRORS"}
                    $canceled = $Records | where-object {$_.result.status -eq "CANCELED"}
                    $failed = $Records | where-object {$_.result.status -eq "FAILED"}
                    $skipped = $Records | where-object {$_.result.status -eq "SKIPPED"}
                    $unknown = $Records | where-object {$_.result.status -eq "UNKNOWN"}
                    
                    # CREATE THE REPORT OBJECT
                    $Object = [ordered]@{
                        hostName = $Records[-1].host.name
                        assetId = $Records[-1].asset.id
                        assetName = $Records[-1].asset.name
                        assetType = $Records[-1].asset.type
                        assetSize = $asset.size
                        assetUom = $asset.uom
                        policyName = $Records[-1].protectionPolicy.name
                        ppdmServer = $_
                        totalBackups = $Records.length
                        okBackups = $okBackups.length
                        okWithErrorsBackups = $okWithErrors.length
                        canceledBackups = $canceled.length
                        failedBackups = $failed.length
                        skippedBackups = $skipped.length
                        unknownBackups = $unknown.length
                    }
    
                    $Activities += New-Object -TypeName psobject -Property $Object
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
$Activities | Export-Csv $ReportOutFile -NoTypeInformation