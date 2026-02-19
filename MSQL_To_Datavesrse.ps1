# Version              Name                                Repository           Description
# -------              ----                                ----------           -----------
# 16.0.25012.12000     Microsoft.Online.SharePoint.PowerS… PSGallery            Microsoft SharePoint Online Services Mod…
# 2.8.19               Microsoft.Xrm.Data.Powershell       PSGallery            This module applies many helpful functio…
# 2.12.0               PnP.PowerShell                      PSGallery            Microsoft 365 Patterns and Practices Pow…
# 22.3.0               SQLServer                           PSGallery            This module allows SQL Server developers…


# the Xrm.Tooling currently is not supported in PowerShell 7.x or above, please use PowerShell 5.x.x (default with PowerShell IDE)
# Install-Module Microsoft.Xrm.Tooling.CrmConnector.PowerShell
# Install-Module SQLServer

#Clear all existing errors from PowerShell scripts
$Error.Clear()

$EntityMaps = @()
$jsonPath = "Entity-Structure.json"


#SQL Server Connection Details
$ServerName = "(localdb)\MSSQLLocalDB"
$DatabaseName = "hptrim"
$UserName = ""
$Password = ""
$SqlConnectionString = 'Data Source={0};database={1};User ID={2};Password={3};ApplicationIntent=ReadOnly' -f $ServerName,$DatabaseName,$UserName,$Password

#for information on how to use SQL Server PowerShell connection details see: 'https://learn.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps'

#SQL Server Queries
$SqlTables = "SELECT TOP 5 name, modify_date FROM sys.tables WHERE type_desc = 'USER_TABLE' ORDER BY name"
$SqlFields = "SELECT DB_NAME() as [Database_Name], SCHEMA_NAME(t.schema_id) as [Schema_Name], t.name AS table_name, c.column_id, c.name AS column_name, c.system_type_id, st.name as data_type, c.max_length, c.precision FROM sys.columns AS c INNER JOIN sys.tables AS t ON t.object_id = c.object_id INNER JOIN sys.types as st ON st.system_type_id = c.system_type_id WHERE t.name='{0}' ORDER BY DB_NAME(), SCHEMA_NAME(t.schema_id), t.name, c.column_id"
$SqlDataTypes = "select name as data_type, system_type_id,max_length, precision, scale, is_nullable from sys.types"

#Dataverse Enterprise Applcation
$EnterpriseUserName = $UserName 
$EnterprisePassword = $Password
$EnterpriseAppId = ""

#Dataverse App Registration Connection Setup
#Dataverse App Registration Connection Setup
$TenantId = "{guid}"
$ClientId = "{guid}"
$ClientSecret = "{secret}"
$BaseURL = "https://{tenantname}.crm6.dynamics.com"
$BaseAPI = "$baseUrl/api/data/v9.2"
$SchemaPrefix = "dev_"


#ConnectionType = UserName or AppRegistration (Default)
$ConnectionType="AppRegistration"

$DataverseConnectionString = $null
if ($ConnectionType.ToLower() -eq "username")
{
    # Connect to Dataverse
    $DataverseConnectionString = "AuthType=OAuth;Username={0};Password={1};Url={2};AppId={3};RedirectUri=app_redirect_url;LoginPrompt=Auto" -f $EnterpriseUserName,$EnterprisePassword,$baseUrl,$EnterpriseAppId
    
} 
else 
{
    $DataverseConnectionString = "AuthType=ClientSecret;Url={0};ClientId={1};ClientSecret={2};Authority=https://login.microsoftonline.com/{3}" -f $BaseAPI,$ClientId,$ClientSecret,$TenantId
}

$crmConnection = Get-CrmConnection -ConnectionString $DataverseConnectionString
$crmOrganisation = Connect-CrmOrganization -Connection $crmConnection


#Check for successfull connection to Dataverse
if($Error.Count -gt 0)
{   
    Write-Host "Unable to establish connection to Dynamics 365 "
    Write-Host "Dataverse Connection : {0}" -f $DataverseConnectionString
    Write-Host $Error
    exit 
}

#Check for successfull connection to Database
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection $SqlConnectionString
$SqlConnection.Open()

if($Error.Count -gt 0)
{
    Write-Host "Unable to establish connection to SQL Database"
    Write-Host "Dataverse Connection : {0}" -f $SqlConnectionString
    Write-Host $Error
    exit 
}


$entity = Get-CrmEntity -Name "Contact"

#Retrieve a List of SQL Database Tables
$Tables = Invoke-Sqlcmd -Query $SqlTables -ServerInstance $ServerName -Database $DatabaseName


#Loop through tables and create Entity Maps
foreach ($Table in $Tables)
{
    Write-Host "Creating EntityMetaData for Table: $($Table.name)"
    #Check if table already exists
    $EntityTable = Get-CrmEntity -Name $Table.Name
    #if ($EntityTable -eq $null)
    #{

        $MetaData = @{
            SchemaName = $Table.Name
            DisplayName = $Table.Name
            Description = "table {0} was mirated from {1}" -f $Table.Name, $DatabaseName
            EntitySetName = "$($Table.Name)s"
            PrimaryNameAttribute = "" # Needs to be set via fieldlist 
            OwnershipType = "UserOwned" # Options: UserOwned or OrganizationOwned
        }
        # Map table to Entity 
        $TableMetadata = @{
            Table = $Table.Name
            EntityMetaData = $Metadata   
            Fields = @()
        }

        $EntityMaps += $TableMetadata
     #}
}
        

foreach ($EntityMap in $EntityMaps)
{
    $Fields = Invoke-Sqlcmd -Query $($SqlFields -f $EntityMap.Table) -ServerInstance $ServerName -Database $DatabaseName
    
    foreach($Field in $Fields)
    {

        #Create new field in current $table if it does not exists
        $columnName = $Fields["column_name"]
        $dataType = $Fields["data_type"]
        $maxLength = $Fields["max_length"]
    
        # Swap out later for Map SQL data types to CRM data types
        $crmAttributeType = "String"; $maxLength = 255 
        switch ($dataType) {
            "int"      { $crmAttributeType = "Integer" }
            "varchar"  { $crmAttributeType = "String"; $maxLength = if ($null -eq $maxLength ) { 255 } else { [int]$maxLength } }
            "nvarchar" { $crmAttributeType = "String"; $maxLength = if ($null -eq $maxLength) { 255 } else { [int]$maxLength } }
            "datetime" { $crmAttributeType = "DateTime" }
            "bit"      { $crmAttributeType = "Boolean" }
            default    { $crmAttributeType = "String"; $maxLength = 255 }
        }

        #Map Table Field to Entitiy Field
        $FieldMetadata = @{
            AttributeSchemaName = $Field.column_name
            AttributeDisplayName = $Field.column_name
            AttributeType = $crmAttributeType 
            MaxLength = $maxLength 
        }

        $EntityMap.Fields += $FieldMetadata
    }
}

# Save the modified JSON to the new path
$EntityMaps | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8
Write-Output "JSON file saved to $jsonPath"

#Clean up connections
$sqlConnection.Close()
$crmConnection.Close()