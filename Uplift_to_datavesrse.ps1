# Version              Name                                Repository           Description
# -------              ----                                ----------           -----------
# 16.0.25012.12000     Microsoft.Online.SharePoint.PowerS… PSGallery            Microsoft SharePoint Online Services Mod…
# 2.8.19               Microsoft.Xrm.Data.Powershell       PSGallery            This module applies many helpful functio…
# 2.12.0               PnP.PowerShell                      PSGallery            Microsoft 365 Patterns and Practices Pow…
# 22.3.0               g                           PSGallery            This module allows SQL Server developers…


# Install-Module Microsoft.Xrm.Tooling.CrmConnector.PowerShell
# Install-Module AMSOFTWARE.Crm

# Connect to Dataverse
$connectionString = "AuthType=OAuth;Username=your_username;Password=your_password;Url=https://{orgvalue}.crm.dynamics.com;AppId=your_app_id;RedirectUri=app_redirect_url;LoginPrompt=Auto"
$crmConnection = Get-CrmConnection -ConnectionString $connectionString


# Define the new table
$tableSchemaName = "new_sampletable"
$tableDisplayName = "Sample Table"
$tableDescription = "This is a sample table created using PowerShell."

# Create a new table
$createTable = @{
    SchemaName = $tableSchemaName
    DisplayName = $tableDisplayName
    Description = $tableDescription
    EntitySetName = "sampletables"
    PrimaryNameAttribute = "name"
    OwnershipType = "UserOwned" # Options: UserOwned or OrganizationOwned
}


# Create the table using AMSOFTWARE Crm module
New-CrmEntity -Connection $crmConnection -EntitySchemaName $createTable.SchemaName `
    -DisplayName $createTable.DisplayName `
    -Description $createTable.Description `
    -PrimaryAttribute $createTable.PrimaryNameAttribute `
    -OwnershipType $createTable.OwnershipType `
    -EntitySetName $createTable.EntitySetName

Write-Host "Table created successfully."



# Import the Microsoft.Xrm.Data.PowerShell module
Import-Module Microsoft.Xrm.Data.PowerShell

# Connect to SQL Server and run a Query to get a list of tables and columns including database tables
$ServerName = "(localdb)\MSSQLLocalDB"
$DatabaseName = "hptrim"
$Query = "SELECT TOP 5 DB_NAME() as [Database_Name], SCHEMA_NAME(t.schema_id) as [Schema_Name], t.name AS table_name, c.column_id, c.name AS column_name, c.system_type_id, st.name as data_type, c.max_length, c.precision FROM sys.columns AS c INNER JOIN sys.tables AS t ON t.object_id = c.object_id INNER JOIN sys.types as st ON st.system_type_id = c.system_type_id ORDER BY DB_NAME(), SCHEMA_NAME(t.schema_id), t.name, c.column_id"

# Load Database DataType mapped to Dataverse Field Types
# Create data_type mapping table between SQL Server and Datavers 
# How to retrieve data type from SQL server
$sql_datatypes = "select name as data_type, system_type_id,max_length, precision, scale, is_nullable from sys.types"
#$DataTypes = Import-Excel "C:\Power\PowerShell\UpliftDatabase\DataType_map.xlsx"

# Connect to the CRM/Dynamics 365 instance, need to update to client/secret
$crmConn = Get-CrmConnection -InteractiveMode
# Invoke-Sqlcmd -Query "SELECT * FROM TSRECORD" -ServerInstance $ServerName -Database $DatabaseName

$reader = Invoke-Sqlcmd -Query $Query -ServerInstance $ServerName -Database $DatabaseName
#$json = $data | ConvertTo-Json
#$json
$existingTable = ""

foreach ($record in $data)
{
    #$publisherPrefix= "crm15"
    $table = $record.Database_Name

    #Create new field in current $table if it does not exists
    $columnName = $reader["colimn_name"]
    $dataType = $reader["data_type"]
    $maxLength = $reader["max_length"]

    Write-Host "Creating attribute for column: $table $columnName, Type: $dataType"

    if ($existingTable -ne $table)
    {
        # Check if the entity exists by querying the metadata
        $entityMetadata = Get-CrmEntity -conn $crmConn -EntityLogicalName $table -ErrorAction SilentlyContinue

        if ($null -eq $entityMetadata) {
            # If the entity does not exist, create a new one
            Write-Host "Entity '$table' does not exist. Creating a new entity."
        
            $displayName = $table 
            $primaryField = "${table} Primary Field"
        
            # Create a new entity using New-CrmEntity cmdlet            
            New-CrmEntity -conn $crmConn -table $table `
                          -EntityDisplayName $displayName `
                          -EntityCollectionName "${table}s" `
                          -PrimaryFieldDisplayName $primaryField `
                          -PrimaryFieldSchemaName "new_primaryfield" `
                          -PrimaryFieldMaxLength 100
                          
            Write-Host "Entity '$table' has been created successfully."
        } else {
            Write-Host "Entity '$table' already exists."
        }

        $existingTable = $table
        continue
    }

    # Swap out later for Map SQL data types to CRM data types
    $crmAttributeType = "String"; $maxLength = 255 
    switch ($dataType) {
        "int" { $crmAttributeType = "Integer" }
        "varchar" { $crmAttributeType = "String"; $maxLength = if ($null -eq $maxLength ) { 100 } else { [int]$maxLength } }
        "nvarchar" { $crmAttributeType = "String"; $maxLength = if ($null -eq $maxLength) { 100 } else { [int]$maxLength } }
        "datetime" { $crmAttributeType = "DateTime" }
        "bit" { $crmAttributeType = "Boolean" }
        default { $crmAttributeType = "String"; $maxLength = 100 }
    }

    # Create a new attribute in CRM for each SQL column
    New-CrmAttribute -conn $crmConn `
                     -table $table `
                     -AttributeSchemaName "new_$columnName" `
                     -AttributeDisplayName "$columnName" `
                     -AttributeType $crmAttributeType `
                     -MaxLength $maxLength -ErrorAction SilentlyContinue    
}


function GetDatabaseConnection {
    param (
        [String] $DatabaseConnectionString
    )    
}

function GetDataverseConnection {
    param (
        [String] $DatavesreConnectionString
    )    
}


function GetTables {
    param (
        [String] $DatabaseName
    )
    
}


function GetFields {
    param (
        [String] $TableName
    )
    
}

function CreateEntityTable {
    param (
        [String] $TableName
    )
    
}

function CreateEntityTableFields {
    param (
        [String] $TableName,
        [String] $Fields
    )
    
}

function RemapFieldTypes {
    param (
        [String] $FieldType,
        [String] $MaxLength

    )

    
}

function GetChangedData {
    param (
        [String] $TableName,
        [String] $Fields,
        [String] $LastRunDate = $null
    )
}


function UpdateEntityTable {
    param (
        [string] $table
    )
    
}

