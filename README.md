# PowerProtect Reports (v19.16)
Pull reports form your PowerProtect Data Manager server(s) that output to excel, or pdf format.

# Dependencies
- Windows 10 or 11
- PowerShell 7.(latest)
- Microsoft® Excel® for Microsoft 365

> [!WARNING]
> Some of these reports can potentially be very large. Ensure you adjusting an parameters within reasonable ranges.

# Reports
| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-all    | All, asset level, protection activities in the last {x} days, including protection storage\storage unit and replication target\storage unit            | xlsx, pdf |

![dm-activities-all](/Assets/dm-activities-all.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-failed | All, asset level, failed protection activities in the last {x} days including the error code, error reason and extended error reason                   | xlsx, pdf |

![dm-activities-failed](/Assets/dm-activities-failed.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-stats  | All, asset level, protection activity status in the last {x} days including assetSize, preCompSize, postCompSize, dedupeRatio, and reductionPercentage | xlsx, pdf |

![dm-activities-stats](/Assets/dm-activities-stats.png)

| Name                  | Description                                                                                                                                           | Output    |
|:---------------------:|:------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-summary | All, asset level, protection activity status in the last {x} summarized by activity status                                                            | xlsx, pdf |

![dm-activities-summary](/Assets/dm-activities-summary.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-assets-nocopies   | All assets with a lastAvailableCopyTime eq null which may indicate a gap in protection                                                                 | xlsx, pdf |

![dm-assets-nocopies](/Assets/dm-assets-nocopies.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-assets-nopolicy   | All assets with a protectionPolicyId eq null which may indicate a gap in protection                                                                    | xlsx, pdf |

![dm-assets-nopolicy](/Assets/dm-assets-nopolicy.png)

 Name                  | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-audit-logs        | All audit audit log entires in the last {x} days                                                                                                       | xlsx, pdf |

![dm-audit-logs](/Assets/dm-audit-logs.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-copies-location   | Get all of the copies for asset type {x} in location {x} (LOCAL, or CLOUD)                                                                             | xlsx, pdf |

![dm-copies-location](/Assets/dm-copies-location.png)


| Name                    | Description                                                                                                                                         | Output    |
|:-----------------------:|:----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-credentials-external | Get all of the external credentials configured within PowerProtect Data manager                                                                     | xlsx, pdf |

![dm-credentials-external](/Assets/dm-credentials-external.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-identities-access | All, identity access account and groups configured with access to PowerProtect Data Manager                                                            | xlsx, pdf |

![dm-identities-access](/Assets/dm-identities-access.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-nas-file          | A list of files protected on your NAS array (requires the PowerProtect Data Manager search engine)                                                     | xlsx, pdf |

![dm-nas-file](/Assets/dm-nas-file.png)

| Name                 | Description                                                                                                                                            | Output    |
|:--------------------:|:-------------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-tsdm-vibs         | vMware installation bundle (VIB) details for PowerProtect Data Manager's transparent data snapshots                                                    | xlsx, pdf |

![dm-nas-file](/Assets/dm-tsdm-vibs.png)

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
| logoScale            | Reduce the logo.png to {x} percent of its original size                                           | decimal |
| reports              | Report specific configurations, these can be left at the default settings                         | array   |
| reports.file         | Name of the PowerShell 7 script being run, used to look up report settings                        | string  |
| reports.reportName   | The name of the report can be set with this property yyyyMMdd-reportName                          | string  |
| reports.tableStyle   | The style of the table you'd like to see the report rednered with in excel                        | string  |
| reports.numberOfDays | The number of days you'd like to return if the filter contains a data parameter                   | int     |
| reports.pdfScale     | Reduce the table to {x} percent of its original size for rednering in pdf format                  | int     |

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
            "file":"dm-activities-stats.ps1",
            "reportName":"dm-activities-stats",
            "tableStyle": "TableStyleMedium2",
            "numberOfDays": 1,
            "pdfScale": 31,
            "pdfOrientation": 2
        },
        {
            "file":"dm-activities-summary.ps1",
            "reportName":"dm-activities-summary",
            "tableStyle": "TableStyleMedium2",
            "numberOfDays": 1,
            "pdfScale": 31,
            "pdfOrientation": 2
        },
        {
            "file":"dm-assets-nocopies.ps1",
            "reportName":"dm-assets-nocopies",
            "tableStyle": "TableStyleMedium3",
            "pdfScale": 42,
            "pdfOrientation": 2
        },
        {
            "file":"dm-assets-nopolicy.ps1",
            "reportName":"dm-assets-nopolicy",
            "tableStyle": "TableStyleMedium3",
            "pdfScale": 42,
            "pdfOrientation": 2
        },
        {
            "file":"dm-audit-logs.ps1",
            "reportName":"dm-audit-logs",
            "tableStyle": "TableStyleMedium5",
            "numberOfDays": 1,
            "pdfScale": 50,
            "pdfOrientation": 2
        },
        {
            "file":"dm-copies-location.ps1",
            "reportName":"dm-copies-location",
            "tableStyle": "TableStyleMedium18",
            "assetType": "VMWARE_VIRTUAL_MACHINE",
            "copyLocation":"LOCAL",
            "pdfScale": 50,
            "pdfOrientation": 2
        },
        {
            "file":"dm-credentials-external.ps1",
            "reportName":"dm-credentials-external",
            "tableStyle": "TableStyleMedium2",
            "pdfScale": 55,
            "pdfOrientation": 2
        },
        {
            "file":"dm-identities-access.ps1",
            "reportName":"dm-identities-access",
            "tableStyle": "TableStyleMedium2",
            "pdfScale": 75,
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
            "file":"dm-tsdm-vibs.ps1",
            "reportName":"dm-tsdm-vibs",
            "tableStyle": "TableStyleMedium2",
            "pdfScale": 50,
            "pdfOrientation": 2
        }
    ]
}
```