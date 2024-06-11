[CmdletBinding()]
    param (
        [Parameter( Mandatory=$false)]
        [switch]$PDF
    )
<#
    THIS CODE REQUIRES POWWERSHELL 7.x.(latest)
        https://github.com/PowerShell/powershell/releases

    IMPORT THE EXCEL INTEROP ASSEMBLY
        I HAD TO DROP THIS ASSEMBLY IN THE SCRIPT FOLDER FROM HERE:
        C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF
#>
$dll = "Microsoft.Office.Interop.Excel.dll"
$office = "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF"
$exists = test-path ".\$($dll)" -PathType Leaf
if(!$exists) {
    Write-Host "`n[WARNING]: Copying $($dll) from:`n$($office)" -ForegroundColor Yellow
    Copy-Item -Path "$($office)\$($dll)" -Destination ".\"
}

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
# .NET ASSEMPLY FOR IMAGES
Add-Type -AssemblyName System.Drawing

$exists = test-path ".\misc\configuration.json" -PathType Leaf
if($exists) {
    # GET THE DEPLOYMENT CONFIGURATION
    $script = $MyInvocation.MyCommand.Name
    $config = get-content ".\misc\configuration.json" | convertfrom-json
    $settings = $config.reports | where-object {$_.file -eq $script}
    $settings
} else {
    # TERMINATING ERROR IF THE CONFIGURATION DOESN'T EXIST
    throw "[ERROR]: Missing the .\misc\configuration.json"
}

# GLOBAL VARS
$global:ApiVersion = 'v2'
$global:Port = 8443
$global:AuthObject = $null

# VARS
$Servers = $config.servers
$Retires = @(1..$config.retries)
$Seconds = $config.seconds
$PageSize = $config.pageSize

# REPORT OPTIONS
$ReportName = $settings.reportName
$ReportOutPath = $config.reportOutPath
$ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).xlsx"
# WHAT ROW TO START THE DATA ON
$HeaderRow = $config.headerRow
<#
    SCALE THE RENDERED PDF DOWN TO $Zoom
    SO IT WILL FIT WITHDH WISE ON THE PAGE
#>
$Zoom = $settings.pdfScale
# VALUES: PORTRIAIT = 1, LANDSCAPE = 2
$Orientation = $settings.pdfOrientation

# LOGO
$LogoPath = $config.logoPath

# SCALE TO SIZE
$LogoScale = $config.logoScale

<#
    ENUMERATIONS FOR THE TABLE STYLES CAN BE FOUND HERE:
        https://learn.microsoft.com/en-us/javascript/api/excel/excel.builtintablestyle?view=excel-js-preview

    SO SWAP IT OUT FOR ONE YOU LIKE BETTER
#>
$TableName = $ReportName
$TableStyle = $settings.tableStyle

# GET ACTIVITIES BASED ON FILTERS
$Date = (Get-Date).AddDays(-$settings.numberOfDays)
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
    $Report = @()
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

    # LAUNCH EXCEL
$Excel = New-Object -ComObject excel.application 
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]

# SURPRESS THE UI
$Excel.visible = $false

# CREATE A WORKBOOK
$Workbook = $Excel.Workbooks.Add()

# GET A HANDLE ON THE FIRST WORKSHEET
$Worksheet = $Workbook.Worksheets.Item(1)

# ADD A NAME TO THE FIRST WORKSHEET
$Worksheet.Name = $ReportName

# LOGO PROPERTIES
$Logo = New-Object System.Drawing.Bitmap $LogoPath

# ADD IMAGE TO THE FIRST WORKSHEET
$Logo = New-Object System.Drawing.Bitmap $LogoPath
$Worksheet.Shapes.AddPicture("$($LogoPath)",1,0,0,0,$Logo.Width*$LogoScale,$Logo.Height*$LogoScale) `
| Out-Null

# DEFINE THE HEADER ROW (row, column)
$Excel.cells.item($HeaderRow,1) = "row"
$Excel.cells.item($HeaderRow,2) = "hostName"
$Excel.cells.item($HeaderRow,3) = "assetId"
$Excel.cells.item($HeaderRow,4) = "assetName"
$Excel.cells.item($HeaderRow,5) = "assetType"
$Excel.cells.item($HeaderRow,6) = "assetSize"
$Excel.cells.item($HeaderRow,7) = "assetUom"
$Excel.cells.item($HeaderRow,8) = "policyName"
$Excel.cells.item($HeaderRow,9) = "ppdmServer"
$Excel.cells.item($HeaderRow,10) = "totalBackups"
$Excel.cells.item($HeaderRow,11) = "okBackups"
$Excel.cells.item($HeaderRow,12) = "okWithErrorsBackups"
$Excel.cells.item($HeaderRow,13) = "canceledBackups"
$Excel.cells.item($HeaderRow,14) = "failedBackups"
$Excel.cells.item($HeaderRow,15) = "skippedBackups"
$Excel.cells.item($HeaderRow,16) = "unknownBackups"

for($i=0;$i -lt $Report.length; $i++) {

    Write-Progress -Activity "Processing records..." `
    -Status "$($i+1) of $($Report.length) - $([math]::round((($i/$Report.length)*100),2))% " `
    -PercentComplete (($i/$Report.length)*100)
    
    # SET THE ROW OFFSET
    $RowOffSet = $HeaderRow+1+$i
    $Excel.cells.item($RowOffSet,1) = $i+1
    $Excel.cells.item($RowOffSet,2) = $Report[$i].hostName
    $Excel.cells.item($RowOffSet,3) = $Report[$i].assetId
    $Excel.cells.item($RowOffSet,4) = $Report[$i].assetName
    $Excel.cells.item($RowOffSet,5) = $Report[$i].assetType
    $Excel.cells.item($RowOffSet,6) = $Report[$i].assetSize
    $Excel.cells.item($RowOffSet,7) = $Report[$i].assetUom
    $Excel.cells.item($RowOffSet,8) = $Report[$i].policyName
    $Excel.cells.item($RowOffSet,9) = $Report[$i].ppdmServer
    $Excel.cells.item($RowOffSet,10) = $Report[$i].totalBackups
    $Excel.cells.item($RowOffSet,11) = $Report[$i].okBackups
    $Excel.cells.item($RowOffSet,12) = $Report[$i].okWithErrorsBackups
    $Excel.cells.item($RowOffSet,13) = $Report[$i].canceledBackups
    $Excel.cells.item($RowOffSet,14) = $Report[$i].failedBackups
    $Excel.cells.item($RowOffSet,15) = $Report[$i].skippedBackups
    $Excel.cells.item($RowOffSet,16) = $Report[$i].unknownBackups

}

<#
    SET CELLS FOR ALL ROWS TO 1.5 TIMES NORAML SIZE
    SET CELLS FOR ALL ROWS TO VERTICALLY ALIGN CENTER
    SO IT DOESN'T HIDE _ CHARACTERS WHEN EXPORTING TO PDF
#>
$WorksheetRange = $Worksheet.UsedRange
$WorksheetRange.EntireRow.RowHeight = $WorksheetRange.EntireRow.RowHeight * 1.5
$WorksheetRange.EntireRow.VerticalAlignment = `
    [Microsoft.Office.Interop.Excel.XLVAlign]::xlVAlignCenter

# AUTO SIZE COLUMNS
$WorksheetRange.EntireColumn.AutoFit() | Out-Null

# CREATE A TABLE IN EXCEL
$TableObject = $Excel.ActiveSheet.ListObjects.Add(`
    [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,`
    $Worksheet.UsedRange,$null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes `
)
# TABLE NAME & STYLE
$TableObject.Name = $TableName
$TableObject.TableStyle = $TableStyle

# EXPORT TO PDF
if($PDF) {
    # PDF SETTINGS
    $Worksheet.PageSetup.Zoom = $Zoom
    $Worksheet.PageSetup.Orientation = $Orientation
    $ReportOutFile = "$($ReportOutPath)\$((Get-Date).ToString('yyyy-MM-dd'))-$($ReportName).pdf"
    $Worksheet.ExportAsFixedFormat($xlFixedFormat::xlTypePDF,$ReportOutFile)
} else {
    $Workbook.SaveAs($ReportOutFile) 
}

# EXIT EXCEL
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null 
Stop-Process -Name "EXCEL"