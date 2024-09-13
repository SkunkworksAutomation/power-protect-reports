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
$ReportName = "dm-copies-location"
$ReportOutPath = "C:\Reports\csv\output"
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).csv"


# GET COPIES BASED ON FILTERS
$Filters = @(
    "not state in (`"DELETED`", `"SOFT_DELETED`")",
    "and not copyType in (`"SPFILE`", `"CONTROLFILE`")",
    "and location eq `"LOCAL`"",
    "and copyType eq `"FULL`""
)
function connect-dmapi {
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
        $login = @{
            username="$($Credential.username)"
            password="$($Credential.GetNetworkCredential().Password)"
        }
        # LOGON TO THE POWERPROTECT API 
        $Auth = Invoke-RestMethod -Uri "https://$($Server):$($Port)/api/$($ApiVersion)/login" `
                    -Method POST `
                    -ContentType 'application/json' `
                    -Body (ConvertTo-Json $login) `
                    -SkipCertificateCheck
        $Object = @{
            server ="https://$($Server):$($Port)/api/$($ApiVersion)"
            token= @{
                authorization="Bearer $($Auth.access_token)"
            } # END TOKEN
        } # END AUTHOBJ

        $global:AuthObject = $Object
        $global:AuthObject.server

    } # END PROCESS
} # END FUNCTION

function disconnect-dmapi {
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
    } # END PROCESS
} # END FUNCTION

function get-dmassetcopies {
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
        $Join = ($Filters -join ' ') -replace '\s','%20' -replace '"','%22'
        $Endpoint = "copies-query"
        $Body = @(
            "filter=$($Join)"
            "orderby=createTime DESC"
        )

        $Query =  Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)?pageSize=$($PageSize)&page=$($Page)" `
        -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Headers ($AuthObject.token) `
        -Body ($Body -join '&') `
        -SkipCertificateCheck

        # CAPTURE THE RESULTS
        $Results = $Query.content
        
        if($Query.page.totalPages -gt 1) {
            
            # INCREMENT THE PAGE NUMBER
            $Page++

            # PAGE THROUGH THE RESULTS
            do {
                $Paging = Invoke-RestMethod -Uri "$($AuthObject.server)/$($Endpoint)?pageSize=$($PageSize)&page=$($Page)" `
                -Method POST `
                -ContentType 'application/x-www-form-urlencoded' `
                -Headers ($AuthObject.token) `
                -Body ($Body -join '&') `
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
function Convert-BytesToSize {
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
    
} # END FUNCTION

$Report = @()
$Servers | ForEach-Object { 
    foreach($Retry in $Retires) {
        try {
            # CONNECT THE THE REST API
            connect-dmapi -Server $_
            Write-Host "[PowerProtect Data Manager]: Getting all FULL copies on LOCAL" `
            -ForegroundColor Green
            
            # QUERY FOR THE COPIES
            $Query = get-dmassetcopies -Filters $Filters -PageSize $PageSize
            
            foreach($Record in $Query) {
           
                $Size = Convert-BytesToSize -Size $Record.size
                $Object = [ordered]@{
                    assetName = $Record.assetName
                    assetId = $Record.assetId
                    assetType = $Record.assetType
                    assetSubtype = $Record.assetSubtype
                    location = $Record.location
                    ppdmServer = $_
                    size = $Size.size
                    uom = $Size.uom
                    createTime = $Record.createTime
                    retentionTime = $Record.retentionTime
                    copyType = $Record.copyType
                    copyConsistency = $Record.copyConsistency
                    retentionLock = $Record.retentionLock
                    retentionLockMode = $Record.retentionLockMode
                }

                $Report += (New-Object -TypeName psobject -Property $Object)
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