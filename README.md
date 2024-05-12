# PowerProtect Reports (v19.16)
Pull reports form your PowerProtect Data Manager server(s) that output to excel, or pdf format.

# Dependencies
- Windows
- PowerShell 7.(latest)

# Usage
- misc\logo.png change this to your logo.png
- misc\configuration.json change this to align with your environment

| Property             | Description                                                                                       | Type    |
|:--------------------:|:--------------------------------------------------------------------------------------------------|:-------:|
| servers              | PowerProtect Data Manager Servers we want to query                                                | array   |
| retries              | The number of retries, in seconds, when trying to connect to a PowerProtect Data Manager server   | int     |
| pagesize             | Size of the pages to be returned by PowerProtect Data Manager                                     | int     |
| reportOutPath        | The system path you want the report files dropped                                                 | string  |
| headerRow            | The number of rows to skip for your logo in the report                                            | int     |
| logoPath             | The system path where your logo is located                                                        | string  |
| logoScale            | The scale you want you logo reduced by                                                            | decimal |
| reports              | Report specific configurations, these can be left at the default settings                         | array   |
| reports.file         | Name of the PowerShell 7 script being run, used to look up report settings                        | string  |
| reports.reportName   | The name of the report can be set with this property yyyyMMdd-reportName                          | string  |
| reports.tableStyle   | The style of the table you'd like to see the report rednered with in excel                        | string  |
| reports.numberOfDays | The number of days you'd like to return if the filter contains a data parameter                   | int     |
| reports.pdfScale     | Scale the data table up or down to a percent of its original size for rendering in pdf format     | int     |

> [!IMPORTANT]
> Reports, with lots of columns, need to be scaled down significantly when rendering to pdf. You can also remove unnecessary column instead.

## Sample configuration.json
```
{
    "servers": [
        "10.x.x.x"
    ],
    "retries": 3,
    "seconds": 10,
    "pagesize": 100,
    "reportOutPath":"C:\\Reports\\output",
    "headerRow": 5,
    "logoPath":"C:\\Reports\\misc\\logo.png",
    "logoScale": 0.18,
    "reports":[
        {
            "file":"dm-activities-all.ps1",
            "reportName":"dm-activities-all",
            "tableStyle": "TableStyleMedium2",
            "numberOfDays": 1,
            "pdfScale": 28,
            "pdfOrientation": 2
        },
        {
            "file":"dm-activities-failed.ps1",
            "reportName":"dm-activities-failed",
            "tableStyle": "TableStyleMedium3",
            "numberOfDays": 1,
            "pdfScale": 16,
            "pdfOrientation": 2
        },
        {
            "file":"dm-identities-access.ps1",
            "reportName":"dm-identities-access",
            "tableStyle": "TableStyleMedium2",
            "pdfScale": 95,
            "pdfOrientation": 2
        },
        {
            "file":"dm-nas-file.ps1",
            "reportName":"dm-nas-file",
            "tableStyle": "TableStyleMedium2",
            "pdfScale": 35,
            "pdfOrientation": 2
        },
        {
            "file":"dm-activities-stats.ps1",
            "reportName":"dm-activities-stats",
            "tableStyle": "TableStyleMedium2",
            "numberOfDays": 1,
            "pdfScale": 31,
            "pdfOrientation": 2
        }
    ]
}
```
 
# Reports
| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-all    | All, asset level, protection activities in the last x days, including protection storage\storage unit and replication target\storage unit            | xlsx, pdf |

![dm-activities-all](/Assets/dm-activities-all.png)

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-failed | All, asset level, failed protection activities in the last x days including the error code, error reason and extended error reason                   | xlsx, pdf |

![dm-activities-failed](/Assets/dm-activities-failed.png)

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-stats  | All, asset level, protection activity status in the last x days including assetSize, preCompSize, postCompSize, dedupeRatio, and reductionPercentage | xlsx, pdf |

![dm-activities-stats](/Assets/dm-activities-stats.png)

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-identities-access | All, identity access account and groups configured with access to PowerProtect Data Manager                                                          | xlsx, pdf |

![dm-identities-access](/Assets/dm-identities-access.png)

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-nas-file          | A list of files protected on your NAS array (requires the PowerProtect Data Manager search engine)                                                   | xlsx, pdf |

![dm-nas-file](/Assets/dm-nas-file.png)