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
cls

$EntityMaps = @()
$jsonPath = "Entity-Structure.json"

#Dataverse Enterprise Applcation
$EnterpriseUserName = $UserName 
$EnterprisePassword = $Password
$EnterpriseAppId = ""

#Dataverse App Registration Connection Setup
$TenantId = "{guid}"
$ClientId = "{guid}"
$ClientSecret = "{secret}"
$BaseURL = "https://{tenantname}.crm6.dynamics.com"
$BaseAPI = "$baseUrl/api/data/v9.2"
$SchemaPrefix = "dev_"
$SolutionName = "sql_migration"

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

# Load the JSON file
$jsonData = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

# Function to create a table (entity)
function Create-DataverseTable {
    param (
        [string]$schemaPrefix,
        [string]$schemaName,
        [string]$displayName,
        [string]$description,
        [string]$primaryNameAttribute,
        [string]$ownershipType
    )

    #$entityRequest = @{
    #    SchemaName = "$($SchemaPrefix)$($schemaName)"
    #    DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, 1033)
    #    Description = New-Object Microsoft.Xrm.Sdk.Label($description, 1033)
    #    OwnershipType = $ownershipType
    #    PrimaryNameAttribute = $primaryNameAttribute
    #}


    #New-CrmEntity -EntityMetadata $entityRequest -Connection $crmConn

    $primaryAttribute = new-object [Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata]
        $primaryAttribute.SchemaName = "$($schemaPrefix)$($primaryNameAttribute)"
        #$primaryAttribute.RequiredLevel = new-object Microsoft.Xrm.Sdk.AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
        #$primaryAttribute.FormatName = FormatName = StringFormatName.Text
        $primaryAttribute.MaxLength=100
        $primaryAttribute.DisplayName = new-object Microsoft.Xrm.Sdk.Label($primaryNameAttribute,1033)
        #$primaryAttribute.Description = "Migrated table"
               
    # Define the entity metadata
    $entity = New-Object Microsoft.Xrm.Sdk.Metadata.EntityMetadata
        $entity.SchemaName = "$($schemaPrefix)$($schemaName)".ToLower()    
        $entity.LogicalName = "$($schemaPrefix)$($schemaName)".ToLower()    
        $entity.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, 1033)  # Display name        
        $entity.DisplayCollectionName  = New-Object Microsoft.Xrm.Sdk.Label("$($displayName)s", 1033)  # Description
        #$entity.OwnershipType = "UserOwned"
        $entity.IsActivity = $false
        #$entity.ChangeTrackingEnabled = $true

    #$entity.OwnerIdType =8 # $EntityMetadata.OwnershipType  # Specify user-owned entity
    #$entity.Description = New-Object Microsoft.Xrm.Sdk.Label($description, 1033)  # Description
    #$entity.EntitySetName = "$($schemaName)s".ToLower()  # Set plural name
    #$entity.IsActivity = $false  # Not an activity entity

    
    # Create a request to create the entity
    $createEntityRequest = New-Object Microsoft.Xrm.Sdk.Messages.CreateEntityRequest
    $createEntityRequest.Entity = $entity
    $createEntityRequest.Parameters.Remove("HasFeedback")
    #$createEntityRequest.PrimaryAttribute = $primaryAttribute       

    Publish-CrmEntity "$($SchemaPrefix)$($schemaName)" -conn $CrmConnection 
      
    return $createEntityRequest
}

# Function to create an attribute (field)
function Create-DataverseField {
    param (
        [string]$entitySchemaPrefix,
        [string]$entitySchemaName,
        [string]$fieldSchemaName,
        [string]$displayName,
        [string]$attributeType,
        [int]$precision = 0,
        [int]$maxLength = 100,
        [bool]$isRequired = $false
    )

    $required = [Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevel]::None 
    #Check if required
    if($isRequired) 
    { 
        [Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevel]::ApplicationRequired 
    }

   
    # Define attribute metadata
    $attributeRequest = @{
        SchemaName = "$($entitySchemaPrefix)$($fieldSchemaName)"
        DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, 1033)
        RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty( $required )
    }
    
    # Set additional properties based on type
    switch ($attributeType) {
        "File"  {}
        "Memo"  {
            $attributeRequest += @{
                #AttributeType = "Memo"
                AttributeTypeCode = "7"
                MaxLength = $maxLength
                Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::TextArea
            }}
        "UniqueIdentifer" {
            $attributeRequest += @{
                #AttributeType = "UniqueIdentifier"
                AttributeTypeCode = "15"
            }
        }        
        "FloatingPoint" {
            $attributeRequest += @{
                #AttributeType = "Double"
                AttributeTypeCode = "4"
                Precision = $precision
                MaxValue = 1000000
                MinValue = 0
            }
        }
        "String" {
            $attributeRequest += @{
                #AttributeType = "String"
                AttributeTypeCode = "14"
                MaxLength = $maxLength
            }
        }
        "WholeNumber" {
            $attribiteField = new-object Microsoft.Xrm.Sdk.Metadata.BigIntAttributeMetadata
            $attribiteField.SchemaName = "$($entitySchemaPrefix)$($fieldSchemaName)"
            $attribiteField.DisplayName = New-Object Microsoft.Xrm.Sdk.Label($displayName, 1033)
            $attribiteField.RequiredLevel = New-Object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty( $required )

            return $attribiteField

            $attributeRequest += @{
                #AttributeType = "WholeNumber"
                AttributeTypeCode = [Microsoft.Xrm.Sdk.Metadata.AttributeTypeCode]::BigInt
                
            }
        }
        "Decimal" {
            $attributeRequest += @{
                #AttributeType = "Decimal"
                AttributeTypeCode = "3"
                Precision = $precision
                MaxValue = 1000000
                MinValue = 0
            }
        }
        "Boolean" {
            $attributeRequest += @{
                #AttributeType = "Boolean"
                AttributeTypeCode = "0"
                OptionSet = New-Object Microsoft.Xrm.Sdk.Metadata.BooleanOptionSetMetadata(
                    (New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadata((New-Object Microsoft.Xrm.Sdk.Label("True", 1033)), 1)),
                    (New-Object Microsoft.Xrm.Sdk.Metadata.OptionMetadata((New-Object Microsoft.Xrm.Sdk.Label("False", 1033)), 0))
                )
            }
        }
        # Add other attribute types as needed
        default {
            Write-Host "Attribute type $attributeType is not supported."
            return
        }

    }

    # Create the request to create the field
    $createAttribute = New-Object Microsoft.Xrm.Sdk.Messages.CreateAttributeRequest
    $createAttribute.EntityName = "$($entitySchemaPrefix)$($entitySchemaName)"
    $createAttribute.Attribute = $attributeRequest 
   
    return $createAttribute
}

# Loop through each entity in the JSON data and create tables and fields
foreach ($entity in $jsonData) {
    $table = $entity.Table
    $schemaName = $entity.EntityMetaData.SchemaName
    $displayName = $entity.EntityMetaData.DisplayName
    $description = $entity.EntityMetaData.Description
    $primaryNameAttribute = $entity.PrimaryNameAttribute
    $ownershipType = $entity.EntityMetaData.OwnershipType

    # Create the table (entity)
    $createEntityRequest = Create-DataverseTable -schemaPrefix $SchemaPrefix -schemaName $schemaName -displayName $displayName -description $description -primaryNameAttribute $primaryNameAttribute -ownershipType $ownershipType

    $CrmConnection.ExecuteCrmOrganizationRequest($createEntityRequest)        
    Publish-CrmEntity "$($SchemaPrefix)$($schemaName)" -conn $CrmConnection    
        Write-Host "Successfully created entity (table): $displayName"


    # Loop through each field and create attributes for the table
    foreach ($field in $entity.Fields) {
        $fieldSchemaName = $field.AttributeSchemaName
        $fieldDisplayName = $field.AttributeDisplayName
        $attributeType = $field.AttributeType
        $precision = $field.AttrubutePrecision
        $maxLength = $field.MaxLength
        $isRequired = $field.AttrubuteRequired

        # Create the field (attribute)
        $createFieldRequest = Create-DataverseField -entitySchemaPrefix $SchemaPrefix -entitySchemaName $schemaName -fieldSchemaName $fieldSchemaName -displayName $fieldDisplayName -attributeType $attributeType -precision $precision -maxLength $maxLength -isRequired $isRequired
    }

    $CrmConnection.ExecuteCrmOrganizationRequest($createEntityRequest)

    Publish-CrmEntity "$($SchemaPrefix)$($schemaName)" -conn $CrmConnection    
}

#Clean up connections
$sqlConnection.Close()
$crmConnection.Dispose()
