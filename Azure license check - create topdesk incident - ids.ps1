# Verbose logging
$verboseLogging = $false

# AzureAD Application Parameters #
$AADtenantID = ""
$AADAppId = ""
$AADAppSecret = ""


# Check the Microsoft docs for the correct sku ids: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# Some commonly used examples:
# Product name                      String ID           GUID
# OFFICE 365 F3	                    DESKLESSPACK	    4b585984-651b-448a-9e53-3b10f069cf7f
# OFFICE 365 E1	                    STANDARDPACK	    18181a46-0d4e-45cd-891e-60aabd171b4e
# Office 365 E3                     ENTERPRISEPACK      6fd2c87f-b296-42f0-b197-1e91e994b900
# Office 365 E5	                    ENTERPRISEPREMIUM	c7df2760-2c81-4ef7-b578-5b5392b571df
# MICROSOFT 365 BUSINESS PREMIUM	SPB	                cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46
# DEVELOPERPACK_E5                  DEVELOPERPACK_E5    c42b9cae-ea4f-4ab7-9717-81576235ccac
# If no value is provided, all licenses will be checked
$licensesToCheck = @(
    @{
        SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"
        ThresholdPercentage = 95
        ThresholdValue = 25
    },
    @{
        SkuId = "c7df2760-2c81-4ef7-b578-5b5392b571df"
        ThresholdPercentage = 100
        ThresholdValue = 10
    }
)

# Default thresholds, only needed when no licensesToCheck are provided
$thresholdPercentageDefault = 90 #Percentage of licenses in use
$thresholdValueDefault = 5 #Amount of licenses left

# Get the license names from the CSV file provided by Microsoft (https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference)
$csvPath = "C:\_Data\Scripting\PowerShell\AzureAD\Product names and service plan identifiers for licensing.csv"
$licenseTableCsv = Import-Csv $csvPath -Delimiter "," -Encoding UTF8
$licenseTableCsvHashTable = @{}
foreach($row in $licenseTableCsv){
    if($licenseTableCsvHashTable.ContainsKey("$($row.'GUID')")){
        # Overwrite Item if already exists
        $licenseTableCsvHashTable."$($row.'GUID')" = $row
        # Or just skip the item alltogether
    }else{
        $licenseTableCsvHashTable.Add("$($row.'GUID')", $row)
    }
}

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try{
    Write-Information -Verbose "Generating Microsoft Graph API Access Token.." -InformationAction Continue
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type      = "client_credentials"
        client_id       = "$AADAppId"
        client_secret   = "$AADAppSecret"
        resource        = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
        # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
        ConsistencyLevel = "eventual";
    }

    Write-Information -Verbose "Searching for licenses.." -InformationAction Continue
    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/subscribedSkus"

    $response = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    if( [String]::IsNullOrEmpty($licensesToCheck) ){
        # Get all licenses of tenant
        $licenses = $response.value
    }else{
        # Get only specified licenses of tenant
        $licenses = $response.value | Where-Object { $_.skuId -in $licensesToCheck.SkuId }
    }
    
    Write-Information -Verbose "Evaluating thresholds for $(@($licenses).Count) license(s).." -InformationAction Continue
    [System.Collections.ArrayList]$licensesWithReachedThreshold = @()
    foreach($license in $licenses){
        Write-Verbose -Verbose:$verboseLogging "Evaluating $($license.skuPartNumber) ($($license.skuId)).." -InformationAction Continue
        # If consumedUnits = 0, don't divide since powershell cannot divide by 0.
        if($license.consumedUnits -eq 0){ 
            $percentageConsumed = 0 
        }else{ 
            $percentageConsumed = ($license.consumedUnits / $license.prepaidUnits.enabled * 100) 
        }
        $availableUnits = $license.prepaidUnits.enabled - $license.consumedUnits

        $licenseName = $licenseTableCsvHashTable["$($license.skuId)"].Product_Display_Name
        if ([String]::IsNullOrEmpty($licenseName)) {
            $licenseName = $license.skuPartNumber
        }

        $licenseToCheckData = $licensesToCheck | Where-Object {$_.SkuId -eq $license.skuId}
        if (-not([String]::IsNullOrEmpty($licenseToCheckData))) {
            $thresholdValue = $licenseToCheckData.thresholdValue
            $thresholdPercentage = $licenseToCheckData.thresholdPercentage
        }
        else{
            $thresholdValue = $thresholdValueDefault
            $thresholdPercentage = $thresholdPercentageDefault
        }

        if( ($percentageConsumed -ge $thresholdPercentage) -or ($availableUnits -le $thresholdValue) ){
            $licenseObject = [PSCustomObject]@{
                name                    =   $licenseName
                skuId                   =   $license.skuId
                thresholdPercentage     =   "$thresholdPercentage%"
                percentageConsumed      =   "$percentageConsumed%"
                thresholdValue          =   $thresholdValue
                availableUnits          =   $availableUnits
            }

            $null = $licensesWithReachedThreshold.Add($licenseObject)
            Write-Verbose -Verbose:$verboseLogging "$licenseName::PercentageConsumed::$($percentageConsumed)%"
            Write-Verbose -Verbose:$verboseLogging "$licenseName::AvailableUnits::$availableUnits"
            Write-Warning -Verbose "Reached threshold for $licenseName ($($license.consumedUnits)/$($license.prepaidUnits.enabled) $($percentageConsumed)%)"
        }
    }
}catch{
    throw "Could not gather licenses from Azure AD. Error: $_"
}


# Create Topdesk incident parameters
$createTopdeskIncident = $true

$TOPdeskUsername = ""
$TOPdeskAPIKey = ""
$TOPdeskBaseUrl = "https://enyoi.topdesk.net/"


$briefDescription = "Office 365 license treshold reached"
$briefDescription = $briefDescription.substring(0, [System.Math]::Min(80, $briefDescription.Length)) # Limit to maximum amount of characters to prevent errors
# Create custom object to display license table
$licensesWithReachedThresholdTopdeskTable = $null
foreach($licenseWithReachedThreshold in $licensesWithReachedThreshold){
    foreach($item in $licenseWithReachedThreshold.PSObject.Properties){
        $licensesWithReachedThresholdTopdeskTable += "<b>$($item.Name)</b> : $($item.Value)<br>"
    }
    $licensesWithReachedThresholdTopdeskTable += "<br>"
}

$request = "
Hi,<br>
<br>
The threshold has been reached for the following Office 365 licenses.<br>
<br>
$($licensesWithReachedThresholdTopdeskTable)<br>
<br>
Kind regards<br>
"

# All the attributes below (except for targetDate should be id's)
$callerLookop =  ""

$entryType = ""
$callType = ""
$category = ""
$subCategory = ""

$branch = ""
$sla = ""

$operatorGroup = ""

$processingStatus = ""

# Format dd-MM-yyyy
#$targetDate = "06-10-2021"

# Accepted values: firstLine, secondLine
$status = "firstLine"

if($createTopdeskIncident -eq $true){
    try {
        # Create authorization headers
        $pair = "${TOPdeskUsername}:${TOPdeskAPIKey}"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $base64 = [System.Convert]::ToBase64String($bytes)
        $key = "Basic $base64"
    
        $headers = @{"authorization" = $Key}
    
        # Make sure base url ends with '/'
        if($TOPdeskBaseUrl.EndsWith("/") -eq $false){
            $TOPdeskBaseUrl = $TOPdeskBaseUrl + "/"
        }
    
        $TOPdeskIncidentObject = @{
            request = "$request"
            briefDescription = "$briefDescription"
            callerLookup = "$callerLookop"
            entryType = "$entryType"
            callType = "$callType"
            category = "$category"
            subcategory = "$subcategory"
            branch = "$branch"
            sla = "$sla"
            operatorGroup = "$operatorGroup"
            processingStatus = "$processingStatus"
        }

        # If Target Date name is specified, format string to required format
        if($targetDate){
            try{
                $datetime = [System.DateTime]::ParseExact($targetDate,"dd-MM-yyyy",$null)
                $null = $TOPdeskIncidentObject.Add( "targetDate", $datetime.ToString("yyyy-MM-ddTHH:mm:ms.fffzz00") )
            }catch{
                Write-Warning -Verbose "The target date [$targetDate] is of an invalid format. The incident will be created without a target date."
            } 
        }
    
        if ($status) {
            $TOPdeskIncidentObject.Add( "status", "$status" )
            Write-Verbose -Verbose:$verboseLogging "Added status '$($status)' to request"
        }
        else {
            Write-Verbose -Verbose:$verboseLogging "No status provided, reverting to default (firstLine)"
        }  
    
        $body = $TOPdeskIncidentObject | ConvertTo-Json -Depth 10
    
        $uriIncidents = $TOPdeskBaseUrl + "tas/api/incidents"
        $incidentResponse = Invoke-RestMethod -Method Post -Uri $uriIncidents -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ContentType 'application/json' -ErrorAction Stop
    
        if($CloseTicket -eq "True"){
            Write-Information -Verbose "TOPdesk incident [$briefDescription] created successfully with number [$($incidentResponse.number)], immediately closed ticket" -InformationAction Continue
        }else{
            Write-Information -Verbose "TOPdesk incident [$briefDescription] created successfully with number [$($incidentResponse.number)]" -InformationAction Continue
        }
    } catch {
        if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
            $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
            Write-Error (("Error creating TOPdesk incident [$briefDescription]. Error: $_" | Out-String) + ($message | Out-String))
        }else{
            Write-Error "Error creating TOPdesk incident [$briefDescription]. Error: $_"
        }
        Write-Verbose -Verbose ("Body: " + ($body | Out-String))
    }    
}