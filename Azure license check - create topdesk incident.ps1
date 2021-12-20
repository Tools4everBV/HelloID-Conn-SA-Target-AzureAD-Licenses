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

# We use the "old" query paraemeters of the TOPdesk API version 1.37.0
# Find the supported attributes at the TOPdesk API documentation: https://developers.topdesk.com/explorer?page=supporting-files&version=1.37.0
$callerUser =  "john.doe@enyoi.org"
$callerAttribute = "email" # Known supported attributes: 'email', 'network_login_name', 'ssp_login_name'
$fallbackUser = "servicedesk@enyoi.org"
$fallbackAttribute = "network_login_name" # Known supported attributes: 'email', 'network_login_name', 'ssp_login_name'

$callerLookupAttribute = 'id' # Attribute to lookup the caller within the incident. Known supported attributes: 'id', 'email'

$briefDescription = "Office 365 license treshold reached"
$briefDescription = New-Object String (,$briefDescription[0..(80-1)]) # Limit to maximum amount of characters to prevent errors
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

$entryTypeName = ""
$callTypeName = "R"
$categoryName = ""
$subCategoryName = ""

$branchName = ""

$impactName = ""
# $urgencyName = ""
# $priorityName = ""

# Mostly the priority decides the duration and target date
# $durationName = ""
# Format dd-MM-yyyy
#$targetDate = "06-10-2021"

$operatorGroupName = ""

$processingStatusName = ""

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
        }
    
        # Get Caller object   
        try{
            $uriCaller = $TOPdeskBaseUrl + "tas/api/persons?$callerAttribute=$callerUser"
            $responseCaller = Invoke-RestMethod -Method Get -Uri $uriCaller -Headers $headers -ContentType 'application/json' -ErrorAction Stop
    
            if( (![String]::IsNullOrEmpty($responseCaller.id)) -and ($responseCaller.count -eq 1) ){
                $caller = $responseCaller
                Write-Verbose -Verbose:$verboseLogging "Caller with $callerAttribute=$callerUser found. ID: $($caller.Id)"
            }elseif([String]::IsNullOrEmpty($responseCaller)){
                Write-Warning -Verbose "Caller with $callerAttribute=$callerUser not found. Trying fallback: $fallbackAttribute=$fallbackUser"
    
                $uriFallBack = $TOPdeskBaseUrl + "tas/api/persons?$fallbackAttribute=$fallbackUser"
                $responseFallback = Invoke-RestMethod -Method Get -Uri $uriFallBack -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $caller = $responseFallback[0]
                
                if($caller.count -eq 1){
                    Write-Verbose -Verbose:$verboseLogging "Fallback user with $fallbackAttribute=$fallbackUser found. ID: $($caller.Id)"
                }else{
                    Write-Error "Fallback user with $fallbackAttribute=$fallbackUser not found."
                }
    
            }elseif($responseCaller.count -gt 1){
                Write-Warning -Verbose "Multiple Callers found with $callerAttribute=$callerUser. Trying fallback: $fallbackAttribute=$fallbackUser"
    
                $uriFallBack = $TOPdeskBaseUrl + "tas/api/persons?$fallbackAttribute=$fallbackUser"
                $responseFallback = Invoke-RestMethod -Method Get -Uri $uriFallBack -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $caller = $responseFallback[0]
    
                Write-Verbose -Verbose:$verboseLogging "Fallback user with $callerfallbackAttributeAttribute=$fallbackUser found. ID: $($caller.Id)"
    
            }
    
            $null = $TOPdeskIncidentObject.Add( "callerLookup", @{$callerLookupAttribute = $caller.$callerLookupAttribute} )
        }catch{
            if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                Write-Error "Could not get TOPdesk Person. Error: $_ $message"
            }else{
                Write-Error "Could not get TOPdesk Person. Error: $_"
            }
        }
    
        # If Entry Type name is specified, get Entry Type object
        if($entryTypeName){
            try{
                $uriEntryTypes = $TOPdeskBaseUrl + "tas/api/incidents/entry_types"
                $responseEntryTypes = Invoke-RestMethod -Method Get -Uri $uriEntryTypes -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $entryTypes = foreach($entryType in $responseEntryTypes){
                    if($entryType.name -eq $entryTypeName){
                        $entryType
                    }
                }
    
                if($null -eq $entryTypes){
                    Write-Warning -Verbose "Entry Type with the name [$entryTypeName] not found, Entry Type is ignored"
                }elseif($entryTypes.count -gt 1){
                    Write-Warning -Verbose "Multiple Entry Types found with the name [$entryTypeName], Entry Type is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "entryType", @{id = $entryTypes.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Warning -Verbose "Could not get TOPdesk Entry Types. Error: $_ $message"
                }else{
                    Write-Warning -Verbose "Could not get TOPdesk Entry Types. Error: $_"
                }
            }
        }
    
    
        # If Call Type name is specified, get Call Type object
        if($callTypeName){
            try{
                $uriCallTypes = $TOPdeskBaseUrl + "tas/api/incidents/call_types"
                $responseCallTypes = Invoke-RestMethod -Method Get -Uri $uriCallTypes -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $callTypes = foreach($callType in $responseCallTypes){
                    if($callType.name -eq $callTypeName){
                        $callType
                    }
                }
    
                if($null -eq $callTypes){
                    Write-Warning -Verbose "Call Type with the name [$callTypeName] not found, Call Type is ignored"
                }elseif($callTypes.count -gt 1){
                    Write-Warning -Verbose "Multiple Call Types found with the name [$callTypeName], Call Type is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "callType", @{id = $callTypes.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Call Types. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Call Types. Error: $_"
                }
            }
        }
    
    
        # If Category name is specified, get Category object
        if($categoryName){
            try{
                $uriCategories = $TOPdeskBaseUrl + "tas/api/incidents/categories"
                $responseCategories = Invoke-RestMethod -Method Get -Uri $uriCategories -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $categories = foreach($category in $responseCategories){
                    if($category.name -eq $categoryName){
                        $category
                    }
                }
    
                if($null -eq $categories){
                    Write-Warning -Verbose "Category with the name [$categoryName] not found, Category is ignored"
                }elseif($categories.count -gt 1){
                    Write-Warning -Verbose "Multiple Categories found with the name [$categoryName], Category is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "category", @{id = $categories.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Categories. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Categories. Error: $_"
                }
            }
        }
    
        # If Sub Category name is specified, get Sub Category object
        if($subCategoryName -and $categoryName){
            try{
                $uriSubCategories = $TOPdeskBaseUrl + "tas/api/incidents/subcategories"
                $responseSubCategories = Invoke-RestMethod -Method Get -Uri $uriSubCategories -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $subCategories = foreach($subCategory in $responseSubCategories){
                    if($subCategory.name -eq $subCategoryName){
                        $subCategory
                    }
                }
    
                if($null -eq $subCategories){
                    Write-Warning -Verbose "Sub Category with the name [$subCategoryName] not found, Sub Category is ignored"
                }elseif($subCategories.count -gt 1){
                    Write-Warning -Verbose "Multiple Sub Categories found with the name [$subCategoryName], Sub Category is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "subcategory", @{id = $subCategories.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Sub Categories. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Sub Categories. Error: $_"
                }
            }
        }
    
        # If Branch name is specified, get branch object
        if($branchName){
            try{
                $uriBranches = $TOPdeskBaseUrl + "tas/api/branches"
                $responseBranches = Invoke-RestMethod -Method Get -Uri $uriBranches -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $branches = foreach($branch in $responseBranches){
                    if($branch.Name -eq $branchName){
                        $branch
                    }
                }
    
                if($null -eq $branches){
                    Write-Warning -Verbose "Branch with the name [$branchName] not found, branch is ignored"
                }elseif($branches.count -gt 1){
                    Write-Warning -Verbose "Multiple Branches found with the name [$branchName], branch is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "branch", @{id = $branches.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk branches. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk branches. Error: $_"
                }
            }
        }
    
        # If Impact name is specified, get branch object
        if($impactName){
            try{
                $uriImpacts = $TOPdeskBaseUrl + "tas/api/incidents/impacts"
                $responseImpacts = Invoke-RestMethod -Method Get -Uri $uriImpacts -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $impacts = foreach($impact in $responseImpacts){
                    if($impact.Name -eq $impactName){
                        $impact
                    }
                }
    
                if($null -eq $impacts){
                    Write-Warning -Verbose "Impact with the name [$impactName] not found, Impact is ignored"
                }elseif($impacts.count -gt 1){
                    Write-Warning -Verbose "Multiple Impacts found with the name [$impactName], Impact is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "impact", @{id = $impacts.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Impacts. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Impacts. Error: $_"
                }
            }
        }
    
    
        # If Urgency name is specified, get branch object
        if($urgencyName){
            try{
                $uriUrgencies = $TOPdeskBaseUrl + "tas/api/incidents/urgencies"
                $responseUrgencies = Invoke-RestMethod -Method Get -Uri $uriUrgencies -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $urgencies = foreach($urgency in $responseUrgencies){
                    if($urgency.Name -eq $urgencyName){
                        $urgency
                    }
                }
    
                if($null -eq $urgencies){
                    Write-Warning -Verbose  "Urgency with the name [$urgencyName] not found, Impact is ignored"
                }elseif($urgencies.count -gt 1){
                    Write-Warning -Verbose "Multiple Urgencies found with the name [$urgencyName], Impact is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "urgency", @{id = $urgencies.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Urgencies. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Urgencies. Error: $_"
                }
            }
        }
    
        # If Priority name is specified, get Priority object
        if($priorityName){
            try{
                $uriPriorities = $TOPdeskBaseUrl + "tas/api/incidents/priorities"
                $responsePriorities = Invoke-RestMethod -Method Get -Uri $uriPriorities -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $priorities = foreach($priority in $responsePriorities){
                    if($priority.Name -eq $priorityName){
                        $priority
                    }
                }
    
                if($null -eq $priorities){
                    Write-Warning -Verbose "Priority with the name [$priorityName] not found, Impact is ignored"
                }elseif($priorities.count -gt 1){
                    Write-Warning -Verbose "Multiple Priorities found with the name [$priorityName], Impact is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "priority", @{id = $priorities.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Priorities. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Priorities. Error: $_"
                }
            }
        }
    
        # If Operator Group name is specified, get Operator Group object
        if($operatorGroupName){
            try{
                [System.Collections.ArrayList]$responseOperatorGroups =  @()
                $take = 100
                $skip = 0
                
                $uriOperatorGroups = $TOPdeskBaseUrl + "tas/api/operatorgroups"  + "?start=$skip&page_size=$take"
                $counter = 0 
                do {
                    if ($counter -gt 0) {
                        $skip += 100
                        $uriOperatorGroups = $TOPdeskBaseUrl + "tas/api/operatorgroups"  + "?start=$skip&page_size=$take"
                    }    
                    $counter++
                    $response = Invoke-WebRequest -Method Get -Uri $uriOperatorGroups -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                    $dataset = $response.content | ConvertFrom-Json
                    foreach($record in $dataset){
                        $null = $responseOperatorGroups.add($record)
                    }
                }until($response.StatusCode -ne '206')
                
                $OperatorGroups = foreach($OperatorGroup in $responseOperatorGroups){
                    if($OperatorGroup.groupName -eq $operatorGroupName){
                        $OperatorGroup
                    }
                }
    
                if($null -eq $OperatorGroups){
                    Write-Warning -Verbose "Operator Group with the name [$operatorGroupName] not found, Operator Group is ignored"
                }elseif($OperatorGroups.count -gt 1){
                    Write-Warning -Verbose "Multiple Operator Groups found with the name [$operatorGroupName], Operator Group is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "operatorGroup", @{id = $OperatorGroups.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Operator Groups. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Operator Groups. Error: $_"
                }
            }
        }
    
        # If Processing Status name is specified, get Processing Status object
        if($processingStatusName){
            try{
                $uriProcessingStatus = $TOPdeskBaseUrl + "tas/api/incidents/processing_status"
                $responseProcessingStatus = Invoke-RestMethod -Method Get -Uri $uriProcessingStatus -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $processingStatuses = foreach($processingStatus in $responseProcessingStatus){
                    if($processingStatus.name -eq $processingStatusName){
                        $processingStatus
                    }
                }
    
                if($null -eq $processingStatuses){
                    Write-Warning -Verbose "Processing Status with the name [$processingStatusName] not found, Processing Status is ignored"
                }elseif($processingStatuses.count -gt 1){
                    Write-Warning -Verbose "Multiple Processing Statuses found with the name [$processingStatusName], Processing Status is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "processingStatus", @{id = $processingStatuses.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Processing Statuses. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Processing Statuses. Error: $_"
                }
            }
        }
    
        # If Processing Status name is specified, get Processing Status object
        if($durationName){
            try{
                $uriDurations = $TOPdeskBaseUrl + "tas/api/incidents/durations"
                $responseDurations = Invoke-RestMethod -Method Get -Uri $uriDurations -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                $durations = foreach($duration in $responseDurations){
                    if($duration.name -eq $durationName){
                        $duration
                    }
                }
    
                if($null -eq $durations){
                    Write-Warning -Verbose "Duration with the name [$durationName] not found, Duration is ignored"
                }elseif($durations.count -gt 1){
                    Write-Warning -Verbose "Multiple Durations found with the name [$durationName], Duration is ignored"            
                }else{
                    $null = $TOPdeskIncidentObject.Add( "duration", @{id = $durations.id} )
                }
            }catch{
                if($_.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
                    $message  = ($_.ErrorDetails.Message | convertFrom-Json).message
                    Write-Error "Could not get TOPdesk Durations. Error: $_ $message"
                }else{
                    Write-Error "Could not get TOPdesk Durations. Error: $_"
                }
            }
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