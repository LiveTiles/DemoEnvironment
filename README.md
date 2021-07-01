# DemoEnvironment
Contains scripts for installing and configuring demo environments.

Instructions for use
1. Create a new M365 tenant, or use an existing one.
2. Note the following information for the M365 tenant
    a. Tenant name
    b. Username and password with SharePoint Administrator permissions
    c. Ensure the above user is also a Term Store Administrator as a new term set group is required. 
4. Clone this repository or download the files as a ZIP
5. Run the DemoInstall script - Install-LiveTilesHub.ps1. Use get-help to see usage information. This will install the LiveTiles Intranet components.
6. Run the DemoImport script - Import-LiveTilesDemoContent.ps1. Use get-help to see usage information. This will import demo SharePoint pages, events, and create a term store group needed for the Workspaces module.
