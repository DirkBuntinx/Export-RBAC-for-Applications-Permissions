# PowerShell-Export-RBAC-for-Applications-Permissions
 PowerShell script to Export RBAC for Applications Permissions from Exchange Online
 
SYNOPSIS
Script to export the RBAC for Applications permissions for each Service Principal in Exchange Online.
IMPORTANT: This is a READ ONLY script; it will only read information from Exchange Online and make no modifications to your tenant.

DESCRIPTION
Script to export the RBAC for Applications permissions for each Service Principal in Exchange Online. 

The script uses 2 mandatory parameters:
---------------------------------------

1) ExportRBACPermissions: This parameter determines which RBAC for Applications Permissions to export, this parameter can have the following values:

    - All: This will export ALL the RBAC for Applications permissions used by both the 'Microsoft Graph' and 'Exchange Web Services (EWS)' APIs for each Service Principal in Exchange Online.

    - EWS: This will export the RBAC for Applications permissions used by the 'Exchange Web Services (EWS)' API for each Service Principal in Exchange Online.

    - Graph: This will export the RBAC for Applications permissions used by the 'Microsoft Graph' API for each Service Principal in Exchange Online.

3) OutputPath: Define the Path to the directory where the output file will be saved. All selected data will be exported to a Tab separated csv file 'Export-RBAC_ApplicationPermissionss_<timestamp>.csv'

                In order to correctly view the files content, the recommendation is to import the data into Excel by using the 'Data Import' method:
                    - Open a Blank Workbook in Excel
                    - Go to the "Data" Tab
                    - Select "Get Data" and select "From File" and click "From Text/csv" and follow the prompts to import the data.

EXAMPLE 1
.\Export-RBACApplicationPermissions_v0.2.ps1 -ExportRBACPermissions:All -OutputPath:'C:\temp'
The will export All the RBAC for Applications permissions to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 2
.\Export-RBACApplicationPermissions_v0.2.ps1 -ExportRBACPermissions:Graph -OutputPath:'C:\temp'
The will export the RBAC for Applications permissions used by the 'Microsoft Graph' API to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory

EXAMPLE 3
.\Export-RBACApplicationPermissions_v0.2.ps1 -ExportRBACPermissions:EWS -OutputPath:'C:\temp'
The will export the RBAC for Applications permissions used by 'Exchange Web Services (EWS)' API to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory 

