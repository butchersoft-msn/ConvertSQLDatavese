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

$CrmConnection = Get-CrmConnection -ConnectionString $DataverseConnectionString
$CrmOrganisation = Connect-CrmOrganization -Connection $CrmConnection


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


# Load the JSON file
$jsonData = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

# Define function to create entities and fields
function Create-DataverseEntityAndFields {
    param (
        $EntityData
    )

    # Extract Entity Metadata    
    $EntityMetaData = $EntityData.EntityMetaData
    $SchemaName = $EntityMetaData.SchemaName
    $DisplayName = $EntityMetaData.DisplayName
    $Description = $EntityMetaData.Description
    $OwnershipType = $EntityMetaData.OwnershipType
    $PrimaryNameAttribute = $EntityData.PrimaryNameAttribute
    $EntitySetName = $EntityMetaData.EntitySetName

    $EntityMetadata = $EntityData.EntityMetaData

    # Get Primary Attribute
    $primaryAttribute = new-object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
    $primaryAttribute.SchemaName = "$($SchemaPrefix)$($EntityData.PrimaryNameAttribute)"
    $primaryAttribute.MaxLength=1
    $primaryAttribute.DisplayName = new-object Microsoft.Xrm.Sdk.Label -ArgumentList @($EntityData.PrimaryNameAttribute,1033)

   
    # Define the entity metadata
    $entity = New-Object Microsoft.Xrm.Sdk.Metadata.EntityMetadata
    $entity.SchemaName = "$($SchemaPrefix)$($EntityMetadata.SchemaName)"    
    $entity.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($EntityMetadata.DisplayName, 1033)  # Display name
    $entity.Description = New-Object Microsoft.Xrm.Sdk.Label($EntityMetadata.Description, 1033)  # Description
    $entity.OwnerIdType =8 # $EntityMetadata.OwnershipType  # Specify user-owned entity
    $entity.EntitySetName = $EntityMetadata.EntitySetName  # Set plural name
    $entity.IsActivity = $false  # Not an activity entity

    
    # Create a request to create the entity
    $createEntityRequest = New-Object Microsoft.Xrm.Sdk.Messages.CreateEntityRequest
    $createEntityRequest.Entity = $entity
    $createEntityRequest.PrimaryAttribute = $primaryAttribute

    $CrmConnection.ExecuteCrmOrganizationRequest($createEntityRequest)

    $EnityFields = $EntityData.Fields
    foreach ($EntityField in $EntityData.Fields) {

        # Define the field (attribute) metadata
        $attributeMetadata = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeMetadata
        $attributeMetadata.SchemaName = "$($SchemaPrefix)$($EnityField.AttributeSchemaName)"  # Schema name of the field
        $attributeMetadata.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($EntityField.AttributeDisplayName, 1033)  # Display name of the field
        $attributeMetadata.AttributeType = $EnityField.AttributeType  # Data type of the field
        

        # Create the request to create the field
        $createAttribute = New-Object Microsoft.Xrm.Sdk.Messages.CreateAttributeRequest
        $createAttribute.EntityName = $EntityData.Table
        $createAttribute.Attribute = $attributeMetadata

        $CrmConnection.ExecuteCrmOrganizationRequest($createAttribute)

    }

    Publish-CrmEntity $entity.SchemaName -conn $CrmConnection    
}



function Map-SqlToDataverseFieldType {
    param (
        [string]$sqlDataType,         # SQL Server data type
        [string]$fieldName,           # Field (attribute) name
        [string]$displayName,         # Display name for the field
        [int]$languageCode = 1033,    # Language code, default is 1033 (English)
        [int]$maxLength = 255,        # Default max length for applicable fields
        [bool]$isRequired = $false,   # Whether the field is required
        [bool]$isAutoIncrement = $false # Whether the field is auto-incremented
    )

    $crmAttributeRequired = $crmAttributeRequired
    $crmAttributeNone = crmAttributeNone 

    # Initialize the attribute metadata based on the SQL data type
    switch ($sqlDataType.ToLower()) {
        "bigint" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.BigIntAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired ) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.AutoNumberFormat = if ($isAutoIncrement) { "{SEQNUM:5}" } else { $null }
        }
        "int" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.IntegerAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.AutoNumberFormat = if($isAutoIncrement) { "{SEQNUM:5}" } else {  $null }
        }
        "decimal" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DecimalAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        "numeric" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DecimalAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        "varchar" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.MaxLength = $maxLength
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::Text
        }
        "nvarchar" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.MaxLength = $maxLength
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::Text
        }
        "text" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.MaxLength = $maxLength
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::TextArea
        }
        "ntext" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.MaxLength = $maxLength
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::TextArea
        }
        "bit" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.BooleanAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.OptionSet = New-Object Microsoft.Xrm.Sdk.Metadata.BooleanOptionSetMetadata(
                (New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadata((New-Object Microsoft.Xrm.Sdk.Label("True", $languageCode)), 1)),
                (New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadata((New-Object Microsoft.Xrm.Sdk.Label("False", $languageCode)), 0))
            )
        }
        "date" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateAndTime
        }
        "datetime" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateAndTime
        }
        "datetime2" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateAndTime
        }
        "smalldatetime" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateAndTime
        }
        "float" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DoubleAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        "real" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.DoubleAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        "uniqueidentifier" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.UniqueIdentifierAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
        }
        "money" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.MoneyAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        "smallmoney" {
            $attribute = New-Object Microsoft.Xrm.Sdk.Metadata.MoneyAttributeMetadata
            $attribute.SchemaName = $fieldName
            $attribute.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, $languageCode)
            $attribute.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(if($isRequired) { $crmAttributeRequired  } else {  crmAttributeNone })
            $attribute.Precision = 2
            $attribute.MaxValue = 1000000
            $attribute.MinValue = 0
        }
        default {
            Write-Error "Unsupported SQL data type: $sqlDataType"
            return $null
        }
    }

    return $attribute
}



# Fun
#Clean up connections
$sqlConnection.Close()
$crmConnection.Dispose()
