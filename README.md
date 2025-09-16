# N2C.SCCM.Migration

Powershell module to assist in SCCM migration.

This module is in early development stage. But it was tested in SCCM Collection migration process.

You could use functions from this script if you are in need.

## Collection Migration use case

Let's say you need to migrate User and Device collection from Old_SCCM to New_SCCM.

First import functions from SCCMCollectionMigration.ps1 at Old_SCCM instance.

```Powershell
. .\SCCMCollectionMigration.ps1
```

Then at you could export all your collection metadata which will be used to recreate collections at New_SCCM like so:

```Powershell
$CollectionsMetadata = Export-xCMCollections
$CollectionsMetadata | Export-Clixml -Path 'C:\collections_export_file.clixml' -Encoding UTF8
```

Using .clixml file at New_SCCM you could recreate collections like so:

```Powershell
Import-CMCollectionsFromMetadata -Path 'C:\collections_export_file.clixml' -SiteCode 'ABC' -ProviderMachineName 'sccm_server.domain.com'
```
