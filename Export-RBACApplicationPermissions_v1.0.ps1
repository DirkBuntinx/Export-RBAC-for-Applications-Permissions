#################################################################################################################################
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. # 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  #
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.               #
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  #
# code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software   #
# product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the  #
# Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims   #
# or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.                 #
#################################################################################################################################
#----------------------------------------------------------------------              
#-     DO NOT CHANGE ANY CODE BELOW THIS LINE                         -
#----------------------------------------------------------------------
#-                                                                    -
#-                           Author:  Dirk Buntinx                    -
#-                           Date:    10/2/2023                       -
#-                           Version: v1.0                            -
#-                                                                    -
#----------------------------------------------------------------------

<#
.SYNOPSIS
Script to export the RBAC for Applications permissions for each Service Principal in Exchange Online.
IMPORTANT: This is a READ ONLY script; it will only read information from Exchange Online and make no modifications to your tenant.

.DESCRIPTION
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

.EXAMPLE
.\Export-RBACApplicationPermissions_v1.0.ps1 -ExportRBACPermissions:All -OutputPath:'C:\temp'
The will export All the RBAC for Applications permissions to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-RBACApplicationPermissions_v1.0.ps1 -ExportRBACPermissions:Graph -OutputPath:'C:\temp'
The will export the RBAC for Applications permissions used by the 'Microsoft Graph' API to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory

.EXAMPLE
.\Export-RBACApplicationPermissions_v1.0.ps1 -ExportRBACPermissions:EWS -OutputPath:'C:\temp'
The will export the RBAC for Applications permissions used by 'Exchange Web Services (EWS)' API to a file called "Export-RBAC_ApplicationPermissions_<timestamp>.csv" in the 'C:\temp' directory 


#>

[CmdletBinding(DefaultParameterSetName ="Default")]
param(

    [Parameter(ParameterSetName="Default",Mandatory=$true, Position=0, HelpMessage="Switch that decides which RBAC permissions to export, values can be: All, Graph or EWS")]
    [ValidateSet('All','Graph', 'EWS')]
    [string]$ExportRBACPermissions,

    [Parameter(ParameterSetName="Default", Mandatory=$true, Position=1, HelpMessage="Define the Path to the directory where the output file will be saved.")]
    [string]$OutputPath=($(Get-Location).Path)
)  

###################################
# Declaring Script wide Variables #
###################################

$Script:Tab = [char]9
$Date = [DateTime]::Now
$Script:StartTime = '{0:MM/dd/yyyy HH:mm:ss}' -f $Date
$Script:FileName = "Export-RBAC_ApplicationPermissions_$('{0:MMddyyyyHHmms}' -f $Date).csv"
$Script:OutputStream = $null
$Script:AllServicePrincipals = @()

$Script:csvOutput = ""
$Script:csvServicePrincipalInfo = ""
$Script:csvMgmtRoleAssigmentInfo = ""
$Script:ProcessRBACPermission = $false
$Script:Graph_RBAC_Roles = @("Application Mail.Read", "Application Mail.ReadBasic", "Application Mail.ReadWrite", "Application Mail.Send", "Application MailboxSettings.Read", 
                            "Application MailboxSettings.ReadWrite", "Application Calendars.Read", "Application Calendars.ReadWrite", "Application Contacts.Read", "Application Contacts.ReadWrite", 
                            "Application Mail Full Access", "Application Exchange Full Access")
$Script:EWS_RBAC_Roles = @("Application EWS.AccessAsApp")


#######################
# BEGIN MAIN FUNCTION #
#######################


Function Export-RBAC_Application_Permissions
{
   
    Begin
    {
        Write-Host "-------------------------------------------"
        Write-Host "- SCRIPT STARTED AT: $($Script:StartTime)  -"
        Write-Host "-------------------------------------------"

        # Call function to Test all the Input parameters and set the required script variables
        Get-InputParameters

        # Call function to Test if the required ExchangeOnline Module v3 is installed
        Test-InstalledModule

        # Call function to create the output file and output stream
        Create-OutputFile
    }

    Process
    {
        # Call function to connect to Exchange Online and get all the ServicePrincipal objects
        Get-EXOServicePrincipals

        $SPIndex = 0
        # Loop through all the ServicePrincipal objects
        Foreach ($ServicePrincipal in $Script:AllServicePrincipals)
        {

            $Script:csvServicePrincipalInfo = ""
            
            Write-Host "ServicePrincipal"
            Write-Host "----------------"
            Write-Host "$($ServicePrincipal.DisplayName)"
            Write-Host "AppId: $($ServicePrincipal.AppId)"
            Write-Host "ID: $($ServicePrincipal.Identity)"
            Write-Host ""

            

            $AllMgmtRoleAssigment = $null   
            # Get all the management scopes that are assigned to this Service Principal
            $AllMgmtRoleAssigment = Get-ManagementRoleAssignment -RoleAssignee $($ServicePrincipal.Identity)
            # First check if any RBAC permissions are assigned, if not we are not saving any data
            if($AllMgmtRoleAssigment -ne $null)
            {
                # Process each ManagementRoleAssigment
                Foreach($MgmtRoleAssigment in $AllMgmtRoleAssigment)
                {
                    
                    Write-Host "$($Script:Tab) MgmtRoleAssigment: $($MgmtRoleAssigment.Name)"
                    Write-Host ""
                    # Check if we are exporting the RBAC permissions or not                
                    $Script:ProcessRBACPermission = $false
                    # based on input parameter ExportRBACPermissions
                    switch ($ExportRBACPermissions) 
                    {
                        'All' 
                            {
                                # All is selected, so we are exporting the RBAC permisisons
                                $Script:ProcessRBACPermission = $true
                            }
                        'EWS' 
                            {
                                # EWS is selected so check if the MgmtRoleAssigment exist in the collection of EWS roles
                                if($($MgmtRoleAssigment.Role) -in $Script:EWS_RBAC_Roles)
                                {
                                    $Script:ProcessRBACPermission = $true
                                }
                            }
                        'Graph'
                            {
                                # Graph is selected so check if the MgmtRoleAssigment exist in the collection of Graph roles
                                if($($MgmtRoleAssigment.Role) -in $Script:Graph_RBAC_Roles)
                                {
                                    $Script:ProcessRBACPermission = $true
                                }
                            }
                        }
                    # end of switch
                    if($Script:ProcessRBACPermission)
                    {
                        # RBAC permissions are assigned for this Service Principal, so we need to save the data
                        # we also increase the index at this point
                        $SPIndex++
                        $Script:csvServicePrincipalInfo = "$SPIndex" + $Script:Tab + $($ServicePrincipal.DisplayName) + $Script:Tab + $($ServicePrincipal.AppId) + $Script:Tab + $($ServicePrincipal.Identity) 
                    
                    
                        $Script:csvMgmtRoleAssigmentInfo = $Script:Tab + $($MgmtRoleAssigment.Name) + $Script:Tab + $($MgmtRoleAssigment.RoleAssigneeType) 
                        #First get the Role details and save the Role Entries to a string
                        $RoleEntriesInfo = ""
                        $firstEntry = $true
                        $ManagementRole = Get-ManagementRole -Identity $($MgmtRoleAssigment.Role)
                        Foreach($RoleEntry in $ManagementRole.RoleEntries)
                        {
                            if($firstEntry)
                            {
                                $RoleEntriesInfo = $RoleEntriesInfo + $($RoleEntry)
                                $firstEntry = $false
                            }else
                            {
                                $RoleEntriesInfo = $RoleEntriesInfo + ", " + $($RoleEntry)
                            }
                        }

                        $Script:csvMgmtRoleAssigmentInfo = $Script:csvMgmtRoleAssigmentInfo + $Script:Tab + $($ManagementRole.Name) + $Script:Tab + $RoleEntriesInfo + $Script:Tab + $($MgmtRoleAssigment.RecipientWriteScope) + $Script:Tab +  $($MgmtRoleAssigment.CustomResourceScope)


                        # Check if the RecipientWriteScope is an Administrative Unit
                        if($MgmtRoleAssigment.RecipientWriteScope -eq 'AdministrativeUnit')
                        {
                            $AdminUnit = Get-AdministrativeUnit -Identity $($MgmtRoleAssigment.CustomResourceScope)
                            $Script:csvMgmtRoleAssigmentInfo = $Script:csvMgmtRoleAssigmentInfo + $Script:Tab + $($AdminUnit.DisplayName) +  $Script:Tab + "N/A"
                        }

                        # Check if the RecipientWriteScope is a Custom Recipient Scope
                        if($MgmtRoleAssigment.RecipientWriteScope -eq 'CustomRecipientScope')
                        {
                            $CustomRecipientScope = Get-ManagementScope -Identity $($MgmtRoleAssigment.CustomResourceScope)
                            $Script:csvMgmtRoleAssigmentInfo = $Script:csvMgmtRoleAssigmentInfo + $Script:Tab + "N/A" + $Script:Tab + $($CustomRecipientScope.RecipientFilter)
                        }
                        $Script:csvOutput = $Script:csvServicePrincipalInfo + $Script:csvMgmtRoleAssigmentInfo
                        Add-Content $Script:OutputStream $Script:csvOutput
                    }
                }
            }else
            {
                Write-Host "$($Script:Tab) No RBAC Application permissions assigned"
                Write-Host ""
            }

        }
        
    }

    End
    {
        $EndTime = '{0:MM/dd/yyyy HH:mm:ss}' -f [DateTime]::Now
        Write-Host "-------------------------------------------"
        Write-Host "- SCRIPT FINISHED AT: $EndTime -"
        Write-Host "-------------------------------------------"
    }
}




########################################
# BEGIN DEFINITION OF HELPER FUNCTIONS #
########################################


# Helper function that validates all the Input Paramters and sets the Script wide variables
Function Get-InputParameters
{
    Write-Host "- Parameters:"
    Write-Host "-------------"
    # Use a switch to display which RBAC Permissions will be exported to the output file (csv)
    # use default value to catch unexpected error
    switch ($ExportRBACPermissions) 
    {
        'All' 
            {
                Write-Host "- Exporting All RBAC for Applications Permissions"
            }
        'EWS' 
            {
                Write-Host "- Exporting Only RBAC for Applications Permissions used by the EWS API"
            }
        'Graph'
            {
                Write-Host "- Exporting Only RBAC for Applications Permissions used by the Graph API"
            }
        default 
            { 
                Write-Host 'Unexpected Error: entering Default for ExportRBACPermissions switch, exiting script' -ForegroundColor Red 
                Exit 
            }
    }
    Write-Host "- Output Directory: $($OutputPath)"
    Write-Host "- Output File Name: $($Script:FileName)"
    Write-Host "-------------------------------------------"
}

# Helper function that validates if the required module is installed
Function Test-InstalledModule
{
    # Test if the required ExchangeOnlineManagement v3 module is installed, if not exit the script and print a help message
    if(get-installedmodule ExchangeOnlineManagement -MinimumVersion 3.3.0) 
    {
        Write-Host "- Module 'ExchangeOnlineManagement' with Minimum version 3.3.0 is installed"
    } 
    else {
        Write-Host "- This script requires 'ExchangeOnlineManagement' module with Minimum version 3.3.0" -ForegroundColor Red
        Write-Host "- Please install the required 'ExchangeOnlineManagement' module from the PSGallery repository by running command:" -ForegroundColor Red
        Write-Host "- install-module -Name ExchangeOnlineManagement -MinimumVersion '3.3.0' -Repository:PSGallery" -ForegroundColor Red
        Exit
               
    }
}

# Helper function that creates the Output csv file and Output stream used to save the data
Function Create-OutputFile
{
    # Create the output file
    # First check if the provided Output Path exists, if not exit the script
    if(!(Test-Path -Path $OutputPath))
    {
        Write-Error "The provided OutputPath does not exist, exiting script" -ForegroundColor Red
        Exit
    }
    else
    {
        # The path exists, so creating the Output file
        $Script:OutputStream = New-Item -Path $OutputPath -Type file -Force -Name $($Script:FileName) -ErrorAction Stop -WarningAction Stop
        # Add the header to the csv file
        $strCSVHeader = "Index" + $Script:Tab + "AppDisplayName" + $Script:Tab + "AppID" + $Script:Tab + "ServicePrincipalID" + $Script:Tab + "MgmtRoleAssigmentName" + $Script:Tab + 
            "RoleAssigneeType" + $Script:Tab + "RoleName" + $Script:Tab + "RoleEntries" + $Script:Tab + "RecipientWriteScope" + $Script:Tab +  "CustomResourceScope" + $Script:Tab + 
            "AdminOrgUnitDisplayName" + $Script:Tab + "RecipientFilter" 
        Add-Content $Script:OutputStream $strCSVHeader
    }
}

# Helper function that connects to Exchange Online and gets the ServicePrincipal objects
Function Get-EXOServicePrincipals
{

    Write-Host "- Connecting to Exchange Online"
    # Connect to Exchange Online
    try
    {        
        Connect-ExchangeOnline -ShowBanner:$false -WarningAction:SilentlyContinue
    }catch [system.exception]
        {
            Write-Host "Error connecting to Exchange Online, exiting script" -ForegroundColor Red
            Write-Host "Command run was:" -ForegroundColor Red
            Write-Host "Connect-ExchangeOnline -ShowBanner:`$false -WarningAction:SilentlyContinue" -ForegroundColor Red
            Exit
        } 

    # Retrieve all the ServicePrincipal objects
    Write-Host "- Retrieving all ServicePrincipal objects from Exchange Online"
    try
    {
        $Script:AllServicePrincipals = Get-ServicePrincipal -WarningAction:SilentlyContinue
    }catch [system.exception]
        {
            Write-Host "Error retrieving all ServicePrincipal objects, exiting script" -ForegroundColor Red
            Write-Host "Command run was:" -ForegroundColor Red
            Write-Host "Get-ServicePrincipal -WarningAction:SilentlyContinue" -ForegroundColor Red
            Exit
        }



    Write-Host "--------------"
    Write-Host "- Found $($Script:AllServicePrincipals.Count) ServicePrincipal objects"
    Write-Host "--------------"
}

########################
# CALL THE MAIN SCRIPT #
########################
Export-RBAC_Application_Permissions