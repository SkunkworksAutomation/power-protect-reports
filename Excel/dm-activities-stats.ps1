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
    "and result.status in (`"OK`",`"OK_WITH_ERRORS`")",
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

    } # END PROCESS
} # END FUNCTION

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
            Write-Host "[PowerProtect Data Manager]: Getting activity stats with a startTime > $($Date.ToString('yyyy-MM-ddThh:mm:ss.fffZ'))" `
            -ForegroundColor Green
            # QUERY FOR THE ACTIVITIES
            $Query = get-dmactivities -Filters $Filters -PageSize $PageSize
            
            foreach($Record in $Query) {
                # BUILD THE CONVERSIONS
                if($null -eq $Record.duration) {
                    $timeSpan = 0
                } else {
                    $timeSpan = New-TimeSpan -Milliseconds $Record.duration
                }
                $assetSize = Convert-BytesToSize -Size $Record.stats.assetSizeInBytes
                $preComp = Convert-BytesToSize -Size $Record.stats.preCompBytes
                $postComp = Convert-BytesToSize -Size $Record.stats.postCompBytes
                $bytesTransferred = Convert-BytesToSize -Size $Record.stats.bytesTransferred

                # BUILD THE REPORT OBJECT
                $Object = [ordered]@{
                    hostName = $Record.host.name
                    assetName = $Record.asset.name
                    assetType = $Record.asset.type
                    assetSize = $assetSize.size
                    assetSizeUom = $assetSize.uom
                    preCompSize = $preComp.size
                    preCompUoM = $preComp.uom
                    postCompSize = $postComp.size
                    postCompUom =  $postComp.uom
                    dedupeRatio = [math]::Round($Record.stats.dedupeRatio,2)
                    reductionPercentage = [math]::Round($Record.stats.reductionPercentage,2)
                    dataTransferred = $bytesTransferred.size
                    dataTransferredUom = $bytesTransferred.uom
                    duration = "{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $timeSpan
                    ppdmServer = $_
                    policyName = $Record.protectionPolicy.name
                    scheduleType = $Record.scheduleInfo.type
                    startTime = $Record.startTime
                    endTime = $Record.endTime
                    nextScheduledTime = $Record.nextScheduledTime
                    jobStatus = $Record.result.status
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
$Excel.cells.item($HeaderRow,3) = "assetName"
$Excel.cells.item($HeaderRow,4) = "assetType"
$Excel.cells.item($HeaderRow,5) = "assetSize"
$Excel.cells.item($HeaderRow,6) = "assetSizeUom"
$Excel.cells.item($HeaderRow,7) = "preCompSize"
$Excel.cells.item($HeaderRow,8) = "preCompUoM"
$Excel.cells.item($HeaderRow,9) = "postCompSize"
$Excel.cells.item($HeaderRow,10) = "postCompUom"
$Excel.cells.item($HeaderRow,11) = "dedupeRatio"
$Excel.cells.item($HeaderRow,12) = "reductionPercentage"
$Excel.cells.item($HeaderRow,13) = "dataTransferred"
$Excel.cells.item($HeaderRow,14) = "dataTransferredUom"
$Excel.cells.item($HeaderRow,15) = "duration"
$Excel.cells.item($HeaderRow,16) = "ppdmServer"
$Excel.cells.item($HeaderRow,17) = "policyName"
$Excel.cells.item($HeaderRow,18) = "scheduleType"
$Excel.cells.item($HeaderRow,19) = "startTime"
$Excel.cells.item($HeaderRow,20) = "endTime"
$Excel.cells.item($HeaderRow,21) = "nextScheduledTime"
$Excel.cells.item($HeaderRow,22) = "jobStatus"

for($i=0;$i -lt $Activities.length; $i++) {

    Write-Progress -Activity "Processing records..." `
    -Status "$($i+1) of $($Activities.length) - $([math]::round((($i/$Activities.length)*100),2))% " `
    -PercentComplete (($i/$Activities.length)*100)
    
    # SET THE ROW OFFSET
    $RowOffSet = $HeaderRow+1+$i
    $Excel.cells.item($RowOffSet,1) = $i+1
    $Excel.cells.item($RowOffSet,2) = $Activities[$i].hostName
    $Excel.cells.item($RowOffSet,3) = $Activities[$i].assetName
    $Excel.cells.item($RowOffSet,4) = $Activities[$i].assetType
    $Excel.cells.item($RowOffSet,5) = $Activities[$i].assetSize
    $Excel.cells.item($RowOffSet,6) = $Activities[$i].assetSizeUom
    $Excel.cells.item($RowOffSet,7) = $Activities[$i].preCompSize
    $Excel.cells.item($RowOffSet,8) = $Activities[$i].preCompUoM
    $Excel.cells.item($RowOffSet,9) = $Activities[$i].postCompSize
    $Excel.cells.item($RowOffSet,10) = $Activities[$i].postCompUom
    $Excel.cells.item($RowOffSet,11) = $Activities[$i].dedupeRatio
    $Excel.cells.item($RowOffSet,12) = $Activities[$i].reductionPercentage
    $Excel.cells.item($RowOffSet,13) = $Activities[$i].dataTransferred   
    $Excel.cells.item($RowOffSet,14) = $Activities[$i].dataTransferredUom
    $Excel.cells.item($RowOffSet,15) = $Activities[$i].duration
    $Excel.cells.item($RowOffSet,16) = $Activities[$i].ppdmServer
    $Excel.cells.item($RowOffSet,17) = $Activities[$i].policyName
    $Excel.cells.item($RowOffSet,18) = $Activities[$i].scheduleType
    $Excel.cells.item($RowOffSet,19) = $Activities[$i].startTime
    $Excel.cells.item($RowOffSet,20) = $Activities[$i].endTime
    $Excel.cells.item($RowOffSet,21) = $Activities[$i].nextScheduledTime
    $Excel.cells.item($RowOffSet,22) = $Activities[$i].jobStatus
    if($Activities[$i].jobStatus -eq "OK_WITH_ERRORS") {
        $Excel.cells.item($RowOffSet,22).Interior.ColorIndex = 44
    }
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