Function Get-GlobalConfig{
    Param(
        [parameter(Mandatory=$true)]
        [STRING]$configFile
    )

    Write-Output "Loading Global Config File"  

    $config = Get-Content $configFile -Raw | ConvertFrom-Json

    # Returning the configuration for later use
    return $config;

} # Configuration loaded!

function get-OathToken{
    Param(
    
        [parameter(Mandatory=$true)]
        [STRING]$resource,
        [parameter(Mandatory=$true)]
        [STRING]$ClientID,
        [parameter(Mandatory=$true)]
        [STRING]$ClientSecret,
        [parameter(Mandatory=$true)]
        [STRING]$loginURL,
        [parameter(Mandatory=$true)]
        [STRING]$tenantdomain             
    )

    # Get an Oauth 2 access token based on client id, secret and tenant domain
    $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body

    #Let's put the oauth token in the header, where it belongs
    $header  = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

    # Returning the header for use in authentication
    return $header
} # Ready for use!

function get-SVCMessages{
    Param(
        [parameter(Mandatory=$true)]
        $ConfigFile,
        [parameter(Mandatory=$false)]
        $OutputPath,
        [parameter(Mandatory=$false)]
        $SPOURL,
        [parameter(Mandatory=$false)]
        $SPOList,
        [parameter(Mandatory=$false)]
        $SPOCreds
    )

    # Retrieving settings
    $globalconfig = Get-GlobalConfig -configFile $ConfigFile

    # Setting us up to get the token
    $ClientID = $globalConfig.AppId
    $ClientSecret = $globalConfig.AppSecret
    $loginURL = $globalConfig.LoginURL
    $tenantdomain = $globalConfig.TenantDomain
    $TenantGUID = $globalConfig.TenantGUID
    $resource = $globalConfig.ResourceAPI

    # Retrieving the token
    $header = get-OathToken -resource $resource -ClientID $ClientID -ClientSecret $ClientSecret -loginURL $loginURL -tenantdomain $tenantdomain

    # Retrieving all messages from the service health dashboard
    $Messages = Invoke-WebRequest -Uri "https://manage.office.com/api/v1.0/$tenantdomain/ServiceComms/Messages" -Headers $header -Method GET | select content
    
    # Getting the content from the raw response, converting that from JSON, and taking only the value (cause that's all we're interested in!)
    $messages = ($Messages.content | ConvertFrom-Json).value
    
    # SharePoint Online List stuff
    if(!([STRING]::IsnullorEmpty($SPOList))){
                   
            if([STRING]::IsnullorEmpty($SPOURL)){
                write-error -Message "SharePoint Online URL not specified! Exiting" -Category ConnectionError -RecommendedAction "Rerun function and specify the 'SPOURL' parameter"
            }

            # Connecting to SPO
            connect-pnponline -url $SPOURL -Credentials $SPOCreds
            
            # seeing if list exists
            $Listcheck = Get-PnPList -Identity $SPOList -ErrorAction SilentlyContinue

            # If the list doesn't exist, create it
            if([STRING]::IsNullOrEmpty($Listcheck)){
                New-PnPList -Title "$SPOList" -Template GenericList -OnQuickLaunch
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Title" -InternalName "SHDTitle"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message ID" -InternalName "SHDID"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Classification" -InternalName "SHDClassification"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Last update" -InternalName "SHDUpdated"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Workload" -InternalName "SHDWorkload"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Feature" -InternalName "SHDFeature"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Affected Tenants" -InternalName "SHDAffectedTenants"
                Add-PnPField -list $SPOList -type Text -DisplayName "Message Affected Users" -InternalName "SHDAffectedUsers"
                Add-PnPField -List $SPOList -Type Note -DisplayName "Message Description" -InternalName "SHDDescription"
            }

            # Retrieving items that already exist
            $Listitems = get-pnplistitem -List $SPOList

    }

    # Now itterate through all if this and write it to an XML file in our chosen output path!
    foreach ($message in $Messages){
        
        #Creating CSV Object
        $CSVObject = New-Object system.collections.arraylist
        $CSVObject = "" | select Title,Status,Severity,WorkloadDisplayName,AffectedTenantCount,AffectedUserCount,Classification,Feature,FeatureDisplayName,ID,ImpactDescription,LastUpdatedTime,MessageType,Message
        
        # Adding information to it
        $CSVObject.Title = $Message.Title
        $CSVObject.Status = $Message.Status
        $CSVObject.Severity = $Message.Severity
        $CSVObject.WorkloadDisplayName = $Message.WorkloadDisplayName
        $CSVObject.AffectedTenantCount = $Message.AffectedTenantCount
        $CSVObject.AffectedUserCount = $Message.AffectedUserCount
        $CSVObject.Classification = $Message.Classification
        $CSVObject.Feature = $Message.Feature
        $CSVObject.FeatureDisplayName = $Message.FeatureDisplayName
        $CSVObject.ID = $Message.ID
        $CSVObject.ImpactDescription = $Message.ImpactDescription
        $CSVObject.LastUpdatedTime = $Message.LastUpdatedTime
        $CSVObject.MessageType = $Message.MessageType
        
        foreach ($i in ($message.messages)){

            $CSVObject.Message += $i.messagetext
        }
        
        if(!([STRING]::IsnullorEmpty($OutputPath))){
            #Setting up filename
            $Filename = $OutputPath + "\" + ($Message.ID) + ".xml"
            $CSVObject | Export-Clixml $Filename
        }

        if(!([STRING]::IsnullorEmpty($SPOList))){
            
            # Transforming properties to variables since add-pnplistitem doesn't like it that way...
            $Title = $Message.Title
            $SHDID = $Message.ID
            $SHDClassification = $Message.Classification
            $SHDUpdated = $Message.LastUpdatedTime
            $SHDWorkload = $Message.WorkloadDisplayName
            $SHDFeature = $Message.Feature
            $SHDAffectedTenants = $Message.AffectedTenantCount
            $SHDAffectedUsers = $Message.AffectedUserCount
            $SHDDescription = $CSVObject.Message

            # See if the item already exists
            [STRING]$ExistentialCheck = $listitems.fieldvalues.Title | Select-String "$SHDID"

            if([STRING]::IsNullOrEmpty($ExistentialCheck)){
                
                Write-Verbose -Message "Adding Item: $SHDID"
                # writing list item
                Add-PnPListItem -List $SPOList -Values @{"Title" = "$SHDID";"SHDTitle"="$Title";"SHDClassification" = "$SHDClassification";"SHDUpdated" = "$SHDUpdated";"SHDWorkload" = "$SHDWorkload";"SHDFeature" = "$SHDFeature";"SHDAffectedTenants" = "$SHDAffectedTenants";"SHDAffectedUsers" = "$SHDAffectedUsers";"SHDDescription" = "$SHDDescription"}
            } Else {
                Write-Verbose -Message "Item already exists in List: $SHDID"
            }
        }
    }

} # AAAAAND we're done!

function get-CurrentStatus{
    Param(
        [parameter(Mandatory=$true)]
        $ConfigFile,
        [parameter(Mandatory=$true)]
        $OutputPath
    )

    # Retrieving settings
    $globalconfig = Get-GlobalConfig -configFile $ConfigFile

    # Setting us up to get the token
    $ClientID = $globalConfig.AppId
    $ClientSecret = $globalConfig.AppSecret
    $loginURL = $globalConfig.LoginURL
    $tenantdomain = $globalConfig.TenantDomain
    $TenantGUID = $globalConfig.TenantGUID
    $resource = $globalConfig.ResourceAPI

    # Retrieving the token
    $header = get-OathToken -resource $resource -ClientID $ClientID -ClientSecret $ClientSecret -loginURL $loginURL -tenantdomain $tenantdomain

    # Retrieving the current status of the service
    $CurrentStatus = Invoke-WebRequest -Uri "https://manage.office.com/api/v1.0/$tenantdomain/ServiceComms/ServiceStatus" -Headers $header -Method GET | select content
    
    # Getting the content from the raw response, converting that from JSON, and taking only the value (cause that's all we're interested in!)
    $CurrentStatus = ($CurrentStatus.content | ConvertFrom-Json).value
    
    # Now itterate through all if this and write it to an XML file in our chosen output path!
    foreach ($i in $CurrentStatus){
        
        #Creating CSV Object
        $CSVObject = New-Object system.collections.arraylist
        $CSVObject = "" | select WorkloadDisplayName,Status,StatusDisplayName,IncidentIds,StatusTime
        
        # Adding information to it
        $CSVObject.WorkloadDisplayName = $i.WorkloadDisplayName
        $CSVObject.Status = $i.Status
        $CSVObject.StatusDisplayName = $i.StatusDisplayName
        
        foreach($incidentID in $i.IncidentIds){
            $CSVObject.IncidentIds += $incidentID +","
        }

        $CSVObject.StatusTime = $i.StatusTime      

        #Setting up filename
        $Filename = $OutputPath + "\CurrentStatus.csv"
        $CSVObject | Export-Csv $Filename -Append -NoTypeInformation
    }

} # AAAAAND we're done!

function get-Services{
    Param(
        [parameter(Mandatory=$true)]
        $ConfigFile,
        [parameter(Mandatory=$true)]
        $OutputPath
    )

    # Retrieving settings
    $globalconfig = Get-GlobalConfig -configFile $ConfigFile

    # Setting us up to get the token
    $ClientID = $globalConfig.AppId
    $ClientSecret = $globalConfig.AppSecret
    $loginURL = $globalConfig.LoginURL
    $tenantdomain = $globalConfig.TenantDomain
    $TenantGUID = $globalConfig.TenantGUID
    $resource = $globalConfig.ResourceAPI

    # Retrieving the token
    $header = get-OathToken -resource $resource -ClientID $ClientID -ClientSecret $ClientSecret -loginURL $loginURL -tenantdomain $tenantdomain

    # Retrieving the current status of the service
    $Services = Invoke-WebRequest -Uri "https://manage.office.com/api/v1.0/$tenantdomain/ServiceComms/Services" -Headers $header -Method GET | select content
    
    # Getting the content from the raw response, converting that from JSON, and taking only the value (cause that's all we're interested in!)
    $Services = ($Services.content | ConvertFrom-Json).value
    
    # Now itterate through all if this and write it to an XML file in our chosen output path!
    foreach ($service in $Services){
        
        #Creating CSV Object
        $CSVObject = New-Object system.collections.arraylist
        $CSVObject = "" | select Id,DisplayName
        
        # Adding information to it
        $CSVObject.Id = $service.Id
        $CSVObject.DisplayName = $service.DisplayName

        #Setting up filename
        $Filename = $OutputPath + "\EnabledServices.csv"
        $CSVObject | Export-Csv $Filename -Append -NoTypeInformation
    }

} # AAAAAND we're done!