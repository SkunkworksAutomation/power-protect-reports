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
        $Endpoint = "$($Endpoint)?parentResourceId=$($Id)"


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
$Excel.cells.item($HeaderRow,2) = "vcenter"
$Excel.cells.item($HeaderRow,3) = "datacenter"
$Excel.cells.item($HeaderRow,4) = "cluster"
$Excel.cells.item($HeaderRow,5) = "esxhost"
$Excel.cells.item($HeaderRow,6) = "esxhostId"
$Excel.cells.item($HeaderRow,7) = "vibResourceId"
$Excel.cells.item($HeaderRow,8) = "vibResourceType"
$Excel.cells.item($HeaderRow,9) = "vibStatus"
$Excel.cells.item($HeaderRow,10) = "vibVersion"
$Excel.cells.item($HeaderRow,11) = "ppdmServer"

for($i=0;$i -lt $Report.length; $i++) {

    Write-Progress -Activity "Processing records..." `
    -Status "$($i+1) of $($Report.length) - $([math]::round((($i/$Report.length)*100),2))% " `
    -PercentComplete (($i/$Report.length)*100)
    
    # SET THE ROW OFFSET
    $RowOffSet = $HeaderRow+1+$i
    $Excel.cells.item($RowOffSet,1) = $i+1
    $Excel.cells.item($RowOffSet,2) = $Report[$i].vcenter
    $Excel.cells.item($RowOffSet,3) = $Report[$i].datacenter
    $Excel.cells.item($RowOffSet,4) = $Report[$i].cluster
    $Excel.cells.item($RowOffSet,5) = $Report[$i].esxhost
    $Excel.cells.item($RowOffSet,6) = $Report[$i].esxhostId
    $Excel.cells.item($RowOffSet,7) = $Report[$i].vibResourceId
    $Excel.cells.item($RowOffSet,8) = $Report[$i].vibResourceType
    $Excel.cells.item($RowOffSet,9) = $Report[$i].vibStatus
    $Excel.cells.item($RowOffSet,10) = $Report[$i].vibVersion
    $Excel.cells.item($RowOffSet,11) = $Report[$i].ppdmServer   

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