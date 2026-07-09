# Scripts

## Publish Power BI repo to SharePoint

### Corporate ELCC SharePoint path

On the corporate device, preview the publish with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\GitHub\Power-BI\scripts\Publish-CorporateSharePoint.ps1" -DryRun
```

You can also double-click `scripts\Preview-CorporateSharePointPublish.cmd` from File Explorer.

Then publish without deleting SharePoint-only files:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\GitHub\Power-BI\scripts\Publish-CorporateSharePoint.ps1"
```

You can also double-click `scripts\Publish-CorporateSharePoint.cmd` from File Explorer.

Use `-Mirror` only when you want the SharePoint folder to exactly match the repo, including deletes.

The corporate wrapper publishes from the repo folder that contains the
`scripts` folder. In the standard checkout that is:

```text
C:\GitHub\Power-BI
```

to:

```text
C:\Users\michael.louie\ESDC EDSC\Federal Secretariat on Early Learning and Child Care - Quants\PBI
```

### Generic usage

Preview the publish first:

```powershell
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -DryRun
```

If PowerShell blocks local scripts on the work device, run it this way:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\Publish-PowerBIToSharePoint.ps1" -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -DryRun
```

Add `-ShowDetails` if you want the full file-by-file preview shown in the
PowerShell window. The detailed output is always written to the log file.

Publish without deleting destination-only files:

```powershell
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI"
```

Publish and make the SharePoint folder match the repo, including deletes:

```powershell
.\scripts\Publish-PowerBIToSharePoint.ps1 -DestinationRoot "C:\Users\you\Org\Site - Documents\Power BI" -Mirror
```

The script excludes Git metadata and local-only Power BI files such as
`localSettings.json` and `cache.abf`. It also skips development tooling folders
such as `scripts` and `skills-for-fabric` unless you add
`-IncludeDevelopmentTools`.
