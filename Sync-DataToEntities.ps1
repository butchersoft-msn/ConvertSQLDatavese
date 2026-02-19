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
$jsonData = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

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

# Connect to SQL Server
$connectionString = "Server=your_server;Database=your_database;Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()

# Load JSON file for entity mapping


# Loop through each table in the JSON
foreach ($tableMapping in $jsonData) {
    $tableName = $tableMapping.Table
    $entityMetaData = $tableMapping.EntityMetaData
    $fieldsMapping = $tableMapping.Fields

    # SQL query to retrieve data from the current table
    $sqlQuery = "SELECT * FROM [$tableName]"
    $command = $connection.CreateCommand()
    $command.CommandText = $sqlQuery
    $dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $dataSet = New-Object System.Data.DataSet
    $dataAdapter.Fill($dataSet)

    # Process each row in the table
    foreach ($row in $dataSet.Tables[0].Rows) {
        # Prepare an object to represent a new Dataverse entity record
        $newEntity = @{}
        
        # Loop through each field in the JSON mapping and populate the entity object
        foreach ($fieldMapping in $fieldsMapping) {
            $sqlColumnName = $fieldMapping.AttributeSchemaName
            $dataverseFieldName = $fieldMapping.AttributeSchemaName
            
            # Check if the SQL column exists in the data and map it
            if ($row.Table.Columns.Contains($sqlColumnName)) {
                $newEntity[$dataverseFieldName] = $row[$sqlColumnName]
            }
        }
        
        # Insert into Dataverse
        # Assuming Connect-CrmOnline has been used to authenticate already
        $entityLogicalName = $entityMetaData.SchemaName
        New-CrmRecord -EntityLogicalName $entityLogicalName -Fields $newEntity
    }

    Write-Output "Migrated data for table $tableName to entity $entityLogicalName."
}

# Close the SQL connection
$connection.Close()
#Clean up connections
$sqlConnection.Close()
$crmConnection.Close()