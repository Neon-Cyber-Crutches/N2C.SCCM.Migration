# N2C.SCCM.Migration

Powershell module to assist in SCCM migration.

This module is in early development stage иut it was tested in real SCCM Collection migration process.

## Install Module

You could install this module from PSGallery

```Powershell
Install-Module N2C.SCCM.Migration
```

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

First import `N2C.SCCM.Migration` module at `Old_SCCM` instance.

```Powershell
Import-Module N2C.SCCM.Migration
```

Then at `Old_SCCM` you could export all your collections metadata which will be used to recreate collections at `New_SCCM` like so:

```Powershell
Export-CMCollectionsMetadata -SiteCode 'ABC' -ProviderMachineName 'sccm_server.domain.com' -Path 'C:\collections_export_file.clixml'
```

Next, using `collections_export_file.clixml` file at `New_SCCM` you could recreate collections like so:

```Powershell
Import-CMCollectionsFromMetadata -Path 'C:\collections_export_file.clixml' -SiteCode 'ABC' -ProviderMachineName 'sccm_server.domain.com'
```
