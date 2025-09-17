# N2C.SCCM.Migration

Powershell module to assist in SCCM migration.

This module is in early development stage. But it was tested in SCCM Collection migration process.

You could use functions from this script if you are in need.

## Implemented Features

- Export User and Device collections metadata to .clixml file
- Recreate User and Device collections preserving
  - Collection name (with optional suffix)
  - Limiting Collection
  - Collection comment
  - Refresh Schedule
  - Membership Rules (all types: Direct, Query, Include, Exclude)
  - Collection location (folders will be created if missing)

## Not Implemented Features

- Migration of Maintenance Windows attached to collections
- Deployments on collections
- Power management on collections
- Alerts on collections
- Security on collection
- Cloud Sync
- Distribution Poin Groups associated with collections
- Collection Variables

## Collection Migration use case

Let's say you need to migrate `User` and `Device` collections from `Old_SCCM` to `New_SCCM`.

First import functions from `SCCMCollectionMigration.ps1` at `Old_SCCM` instance.

```Powershell
. .\SCCMCollectionMigration.ps1
```

Then at `Old_SCCM` you could export all your collections metadata which will be used to recreate collections at `New_SCCM` like so:

```Powershell
$CollectionsMetadata = Export-xCMCollections
$CollectionsMetadata | Export-Clixml -Path 'C:\collections_export_file.clixml' -Encoding UTF8
```

Next using `collections_export_file.clixml` file at `New_SCCM` you could recreate collections like so:

```Powershell
Import-CMCollectionsFromMetadata -Path 'C:\collections_export_file.clixml' -SiteCode 'ABC' -ProviderMachineName 'sccm_server.domain.com'
```
