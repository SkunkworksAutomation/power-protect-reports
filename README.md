# power-protect-reports
Pull reports form your PowerProtect Data Manager server(s) that output to excel, or pdf

# Usage

 
# Reports
| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-all    | All, asset level, protection activities in the last x days, including protection storage\storage unit and replication target\storage unit            | xlsx, pdf |

![dm-activities-all](/Assets/dm-activities-all.png)

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-failed | All, asset level, failed protection activities in the last x days including the error code, error reason and extended error reason                   | xlsx, pdf |

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-activities-stats  | All, asset level, protection activity status in the last x days including assetSize, preCompSize, postCompSize, dedupeRatio, and reductionPercentage | xlsx, pdf |

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-identities-access | All, identity access account and groups configured with access to PowerProtect Data Manager                                                          | xlsx, pdf |

| Name                 | Description                                                                                                                                          | Output    |
|:--------------------:|:-----------------------------------------------------------------------------------------------------------------------------------------------------|:---------:|
| dm-nas-file          | Nas file report                                                                                                                                      | xlsx, pdf |