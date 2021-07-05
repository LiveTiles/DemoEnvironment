# DemoEnvironment
Contains scripts for installing and configuring demo environments.

Instructions for use
1. Create a new M365 tenant, or use an existing one.
2. Note the following information for the M365 tenant
    a. Tenant name
    b. Username and password with SharePoint Administrator permissions
    c. Ensure the above user is also a Term Store Administrator as a new term set group is required. 
3. Clone this repository or download the files as a ZIP
4. Run the DemoInstall script - Install-LiveTilesHub.ps1. Use get-help to see usage information. This will install the LiveTiles Intranet components.
5. Run the DemoImport - Content script - Import-LiveTilesDemoContent.ps1. Use get-help to see usage information. This will import demo SharePoint pages, events, and create a term store group needed for the Workspaces module.
6. Run the DemoImport - Configuration script - Import-LiveTilesConfiguration.ps1. Use get-help to see usage information. This will import demo configuration for LiveTiles Intranet.
7. Manually import LiveTiles Workspaces configuration. From the JsonFiles folder, use the files [TENANT_NAME]-siteType-Community.json, [TENANT_NAME]-siteType-Project.json, and [TENANT_NAME]-siteType-Team.json to provision 3 new site types Community, Project, and Team.
8. Manually import LiveTiles Metadata configuration. From the JsonFiles folder, use the files [TENANT_NAME]-metadata-Department.json, [TENANT_NAME]-metadata-Project.json, and [TENANT_NAME]-metadata-Team.json to provision 3 new metadata configurations for Department, Project, and Team.

On completion, you will see a homepage similar to the following:

![image](https://user-images.githubusercontent.com/17925147/124096963-623d5d80-da5b-11eb-923e-fe0c32dc0b45.png)
