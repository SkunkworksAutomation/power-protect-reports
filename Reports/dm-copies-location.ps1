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

# FILTER VARS
$Type = $settings.assetType
$Location = $settings.copyLocation

# FILTER 
$Filters = @(
    "assetType eq `"$($Type)`"",
    "and not state in (`"DELETED`", `"SOFT_DELETED`")",
    "and not copyType in (`"SPFILE`", `"CONTROLFILE`")",
    "and location eq `"$($Location)`""
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
            Write-Host "[PowerProtect Data Manager]: Getting all $($Type) copies in location: $($Location)" `
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
$Excel.cells.item($HeaderRow,2) = "assetName"
$Excel.cells.item($HeaderRow,3) = "assetId"
$Excel.cells.item($HeaderRow,4) = "assetType"
$Excel.cells.item($HeaderRow,5) = "assetSubtype"
$Excel.cells.item($HeaderRow,6) = "location"
$Excel.cells.item($HeaderRow,7) = "ppdmServer"
$Excel.cells.item($HeaderRow,8) = "size"
$Excel.cells.item($HeaderRow,9) = "uom"
$Excel.cells.item($HeaderRow,10) = "createTime"
$Excel.cells.item($HeaderRow,11) = "retentionTime"
$Excel.cells.item($HeaderRow,12) = "copyType"
$Excel.cells.item($HeaderRow,13) = "copyConsistency"
$Excel.cells.item($HeaderRow,14) = "retentionLock"
$Excel.cells.item($HeaderRow,15) = "retentionLockMode"

for($i=0;$i -lt $Report.length; $i++) {

    Write-Progress -Activity "Processing records..." `
    -Status "$($i+1) of $($Report.length) - $([math]::round((($i/$Report.length)*100),2))% " `
    -PercentComplete (($i/$Report.length)*100)
    
    # SET THE ROW OFFSET
    $RowOffSet = $HeaderRow+1+$i
    $Excel.cells.item($RowOffSet,1) = $i+1
    $Excel.cells.item($RowOffSet,2) = $Report[$i].assetName
    $Excel.cells.item($RowOffSet,3) = $Report[$i].assetId
    $Excel.cells.item($RowOffSet,4) = $Report[$i].assetType
    $Excel.cells.item($RowOffSet,5) = $Report[$i].assetSubtype
    $Excel.cells.item($RowOffSet,6) = $Report[$i].location
    $Excel.cells.item($RowOffSet,7) = $Report[$i].ppdmServer
    $Excel.cells.item($RowOffSet,8) = $Report[$i].size
    $Excel.cells.item($RowOffSet,9) = $Report[$i].uom
    $Excel.cells.item($RowOffSet,10) = $Report[$i].createTime
    $Excel.cells.item($RowOffSet,11) = $Report[$i].retentionTime
    $Excel.cells.item($RowOffSet,12) = $Report[$i].copyType
    $Excel.cells.item($RowOffSet,13) = $Report[$i].copyConsistency
    $Excel.cells.item($RowOffSet,14) = $Report[$i].retentionLock
    $Excel.cells.item($RowOffSet,15) = $Report[$i].retentionLockMode

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