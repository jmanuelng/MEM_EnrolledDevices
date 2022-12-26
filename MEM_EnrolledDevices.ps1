
<#
.SYNOPSIS
    Reports Intune enrolled devices to CSV

.DESCRIPTION
	Get list of devices that have enrolled to MEM in the last X minutes, days, etc.
    - Accepts number of days as parameter
    - Exports results to CSV file. Report get created in current working folder.
    - Can be filtered by AAD Group

.NOTES
    Sources and inspiration:
    https://github.com/microsoftgraph/powershell-intune-samples/blob/master/CheckStatus/Check_enrolledDateTime.ps1
    https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-1.0
    https://social.technet.microsoft.com/Forums/en-US/a61be7a4-5eba-4e42-8544-f632b7f11ae1/case-sharing-powershell-to-query-device-group-membership
    https://github.com/jordanbardwell/Scripts/blob/main/Get-ObjectIdByDeviceId.ps1

    $Device = Get-MgDevice -Filter "deviceId eq '$DeviceId'"

.EXAMPLE
    MEM_EnrolledDevices 365

    Will create a report for all enrolled devices in the las 365 days.

.EXAMPLE
    MEM_EnrolledDevices 30 'MEM Devices Windows'
    
    Creates report for all enrolled devices in the las 30 days that belong to the Azure AD group "MEM Devices Windows"

#>

#Region Parameters

# Reading script parameters
[CmdletBinding()]
param (
    [int]$Days=30,
    [string]$GroupFilter
    )

#Endregion Parameters


#Region Functions


function Get-UriCallError ([System.InvalidOperationException]$Exception, [string]$uri, [string]$errDescription) {

    <#
    .SYNOPSIS
    Function used to cath Graph errors
    .DESCRIPTION
    Catch errors when trying to connect to Microsoft Graph, for troubleshooting purposes, displays friendly description.
    .EXAMPLE
    Get-UriCallError $_.Exception $uri "Error description"
    Displays error with URI path detial and description "Error Description"
    .NOTES
    NAME: Get-UriCallError
    #>

    #@jmanuelnieto: Function to catch an error and display error info.
    $errorResponse = $Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();

    Write-Host "Error Description: $errDescription" -f Red
    Write-Host
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Host
    if (!($uri -eq $null -or $uri -eq "")) {
        Write-Error "Request to $uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    }        
    Write-Host

    break

}


function Get-AuthToken {

    <#
    .SYNOPSIS
    This function is used to authenticate with the Graph API REST interface
    .DESCRIPTION
    The function authenticate with the Graph API Interface with the tenant name
    .EXAMPLE
    Get-AuthToken
    Authenticates you with the Graph API interface
    .NOTES
    NAME: Get-AuthToken
    #>
    
    [cmdletbinding()]
    
    param
    (
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
    $tenant = $userUpn.Host
    
    Write-Host "Checking for AzureAD module..."
            $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    
        if ($null -eq $AadModule) {
                Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
            $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    
        }
    
        if ($null -eq $AadModule) {
            write-host
            write-host "AzureAD Powershell module not installed..." -f Red
            write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
            write-host "Script can't continue..." -f Red
            write-host
            exit
        }
    
    # Getting path to ActiveDirectory Assemblies
    # If the module count is greater than 1 find the latest version
    
        if($AadModule.count -gt 1){
            $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
            $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
    
                # Checking if there are multiple versions of the same module found
                if($AadModule.count -gt 1){
                    $aadModule = $AadModule | Select-Object -Unique
                }
    
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"    
        }
        else {
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    
    $clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$Tenant"
    
        try {

        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    
        # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
        # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
    
        $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"    
        $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
        $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result
    
            # If the accesstoken is valid then create the authentication header
            if($authResult.AccessToken){
    
            # Creating header for Authorization token
            $authHeader = @{
                'Content-Type'='application/json'
                'Authorization'="Bearer " + $authResult.AccessToken
                'ExpiresOn'=$authResult.ExpiresOn
                }
    
            return $authHeader
    
            }
    
            else {
    
            Write-Host
            Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
            Write-Host
            break
    
            }
    
        }
    
        catch {
    
        write-host $_.Exception.Message -f Red
        write-host $_.Exception.ItemName -f Red
        write-host
        break
    
        }
    
}
    
  
function Connect-MsftGraph {

    $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
    $tenant = $userUpn.Host
    $User = ""
    $global:authToken = Get-AuthToken -User $User
    $deviceID = "37c5ba7c-8763-4c2e-87cc-695038ce0950"


    $graphApiVersion = "v1.0"
    $Device_resource = "devices"

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($Device_resource)/$deviceID/memberOF"
    Write-Verbose $uri
    $Groups = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

    $GroupsCount = @($Groups).count

    Write-Host "Son " $GroupsCount "Grupos"


}



Function Get-AADUser(){

<#
.SYNOPSIS
This function is used to get AAD Users from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and gets any users registered with AAD
.EXAMPLE
Get-AADUser
Returns all users registered with Azure AD
.EXAMPLE
Get-AADUser -userPrincipleName user@domain.com
Returns specific user by UserPrincipalName registered with Azure AD
.NOTES
NAME: Get-AADUser
#>

[cmdletbinding()]

param
(
    $userPrincipalName,
    $Property
)

# Defining Variables
$graphApiVersion = "v1.0"
$User_resource = "users"
    
    try {
        
        if($userPrincipalName -eq "" -or $null -eq $userPrincipalName){
        
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)"
        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
        
        }

        else {
            
            if($Property -eq "" -or $null -eq $Property){

                $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$userPrincipalName"
                Write-Verbose $uri
                Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get

            }

            else {

                $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$userPrincipalName/$Property"
                Write-Verbose $uri
                (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

            }

        }
    
    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break

    }

}


function Get-DeviceGroups ($azureDeviceId) {

<#
.SYNOPSIS
Used to query AAD to get list of groups for a specific device
.DESCRIPTION
Connects to Graph API Interface and gets list of all groups for a device specified via Azure Device ID as parameter
#>
    
    if($global:authToken) {

        if (!($null -eq $azureDeviceId -or $azureDeviceId -eq "")) {

            
                $graphApiVersion = "v1.0"
                $Device_resource = "devices"


                $device_uri = "https://graph.microsoft.com/$graphApiVersion/$($Device_resource)?`$filter=deviceID eq '$azureDeviceId'"
                
                try {

                    $Device = (Invoke-RestMethod -Uri $device_uri -Headers $authToken -Method Get).value
                    $deviceId = $Device.id
                }
                
                catch {

                    Get-UriCallError $_.Exception $device_uri "Error while getiing Device, filtered for AzureAD ID: $azureDeviceId `n`t`t`tIn Get-DeviceGroups function"

                }


                if (!($null -eq $deviceId -or $deviceId -eq "")) {
                
                    $deviceGroups_uri = "https://graph.microsoft.com/$graphApiVersion/$($Device_resource)/$deviceId/memberOf"

                    try {

                        $deviceGroups = (Invoke-RestMethod -Uri $deviceGroups_uri -Headers $authToken -Method Get).value

                    }

                    catch {

                        Get-UriCallError $_.Exception $deviceGroups_uri "Getting group membership for device with AzureAD DeviceID:$azureDeviceId `n`t`t`tAzureAD ObjectID for Device: $deviceId `n`t`t`tIn Get-DeviceGroups function"

                    }
                }

                else {
                    
                    $deviceGroups = @()
                }

        }

        else {

            $deviceGroups = @()
        }


        return $deviceGroups

    }

    else {

        $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
        $global:authToken = Get-AuthToken -User $User

        
    }
}



function Get-DeviceInfo ($azureDeviceId) {

<#
.SYNOPSIS
Gets device information for a Device
.DESCRIPTION
Connects to Graph API Interface and gets the following properties for a given device as specified via Azure Device ID as parameter. 
Properties queried:
    - Account enabled
    - Disply name
    - Manufacturer
    - Operating System
    - OS Version
#>
    
    if($global:authToken) {

        if (!($null -eq $azureDeviceId -or $azureDeviceId -eq "")) {

            $graphApiVersion = "v1.0"
            $azDevice_resource = "devices"

            try {

                $azDevice_uri = "https://graph.microsoft.com/$graphApiVersion/$($azDevice_resource)?`$filter=deviceID eq '$azureDeviceId'&`$select=accountEnabled,displayName,manufacturer,operatingSystem,operatingSystemVersion"
                $azDeviceInfo = (Invoke-RestMethod -Uri $azDevice_uri -Headers $authToken -Method Get).value

            }

            catch {
                
                Get-UriCallError $_.Exception $azDevice_uri "Error while trying to fetch information details for Device $azureDeviceId"

            }
            
        }

        else {

            $azDeviceInfo = @()

        }


        return $azDeviceInfo

    }

    else {

        $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
        $global:authToken = Get-AuthToken -User $User

        
    }


}


function Get-DeviceHardware ($DeviceId) {

    <#
    .SYNOPSIS
    Gets device hardware information for a Device
    .DESCRIPTION

    Connects to Graph API Interface and gets the following properties for a given device as specified via Intune Device ID as parameter. 
    Special thanks to Ben Hopper (tw:@BenHopperAU) for guiding me on how to correctly get HW info.
    
    Properties queried:
        - ChassisType
        - OS BuildNumber
        - OS Language
        - IPv4
        - Subnet
        - Cellular Technology
        - IMEI
        - Others
    #>
        
        if($global:authToken) {
    
            if (!($null -eq $DeviceId -or $DeviceId -eq "")) {
    
                $graphApiVersion = "beta"
                $Device_resource = "deviceManagement/managedDevices"
    
                try {
    
                    $Device_uri = "https://graph.microsoft.com/$graphApiVersion/$($Device_resource)/{$DeviceId}?`$select=hardwareInformation"
                    $DeviceHwInfo = (Invoke-RestMethod -Uri $Device_uri -Headers $authToken -Method Get)
    
                }
    
                catch {
                    
                    Get-UriCallError $_.Exception $Device_uri "Error while trying to fetch information details for Device $DeviceId"
    
                }
                
            }
    
            else {
    
                $DeviceHwInfo = @()
    
            }
    
    
            return $DeviceHwInfo
    
        }
    
        else {
    
            $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            $global:authToken = Get-AuthToken -User $User
    
            
        }
    
    
    }



function Get-UserGroups ($azureUserId) {

<#
.SYNOPSIS
Gets list of groups a user belongs to
.DESCRIPTION
Connects to Graph API Interface and gets list of all groups for a the user specified in parameter via Azure User ID
#>
    
    if($global:authToken) {

        if (!($null -eq $azureUserId -or $azureUserId -eq "")) {

            
            $graphApiVersion = "v1.0"
            $user_resource = "users"

            $userFind_uri = "https://graph.microsoft.com/$graphApiVersion/$($user_resource)/$azureUserId"
            $userGroups_uri = "https://graph.microsoft.com/$graphApiVersion/$($user_resource)/$azureUserId/memberOf"

            #Confirm if user exists, not deleted
            try {
                $userFind = (Invoke-RestMethod -Uri $userFind_uri -Headers $authToken -Method Get).id
            }
            catch {
                $userFind = "NOT FOUND"
            }

            try {
                if (!($userFind -eq "NOT FOUND")) {
                    $userGroups = (Invoke-RestMethod -Uri $userGroups_uri -Headers $authToken -Method Get).value
                }
                else {
                    $userGroups = "USER NOT FOUND"
                }
            }

            catch {

                Get-UriCallError $_.Exception $userGroups_uri "Getting group membership for user with AzureAD UserID:$azureUserId `n`t`t`tIn Get-UserGroups function"

            }


        }

        else {

            $userGroups = @()
        }


        return $userGroups

    }

    else {

        $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
        $global:authToken = Get-AuthToken -User $User

        
    }
}



function Get-UserInfo ($azureUserId) {

<#
.SYNOPSIS
Get information for a specified user
.DESCRIPTION
Connects to Graph API Interface and gets properties for the user specified via Azure User ID as parameter
Properties queried:
    - Display name
    - Account enabled
    - Company name
    - Country
    - City
    - Usage location
#>
    
    if($global:authToken) {

        if (!($null -eq $azureUserId -or $azureUserId -eq "")) {

            $graphApiVersion = "v1.0"
            $User_resource = "users"

            $userFind_uri = "https://graph.microsoft.com/$graphApiVersion/$($user_resource)/$azureUserId"
            $user_uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)?`$filter=id eq '$azureUserId'&`$select=displayName,accountEnabled,companyName,country,city,usageLocation"

            #Confirm if user exists, not deleted
            try {
                $userFind = (Invoke-RestMethod -Uri $userFind_uri -Headers $authToken -Method Get).id
            }
            catch {
                $userFind = "NOT FOUND"
            }

            try {

                if (!($userFind -eq "NOT FOUND")) { 
                    $UserInfo = (Invoke-RestMethod -Uri $user_uri -Headers $authToken -Method Get).value
                }
                else {
                    $userInfo = "NOT FOUND"
                }

            }

            catch {
                
                Get-UriCallError $_.Exception $user_uri "Error while trying to fetch User Information for $azureUserId"

            }
            
        }

        else {

            $UserInfo = @()

        }


        return $UserInfo

    }

    else {

        $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
        $global:authToken = Get-AuthToken -User $User

        
    }


}


#Endregion Functions


#Region Authentication
    
write-host

# Checking if authToken exists before running authentication
if($global:authToken){

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

        if($TokenExpires -le 0){

        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host

            # Defining Azure AD tenant name, this is the name of your Azure Active Directory (do not use the verified domain name)    
            if($null -eq $User -or $User -eq ""){    
                $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
                Write-Host
            }

        $global:authToken = Get-AuthToken -User $User

        }
}

# Authentication doesn't exist, calling Get-AuthToken function
    
else {

    if($null -eq $User -or $User -eq ""){

    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
    Write-Host

    }

# Getting the authorization token
$global:authToken = Get-AuthToken -User $User

}

#Endregion Authentication


#Region Main

$msg = "`n`nGathering information from Microsoft Endpoint Manager using Microsoft Graph"
Write-Host $msg -ForegroundColor White

# Filter for the minimum number of minutes when the device enrolled into the Intune Service
# 1440 = 24 hours

$CurrentTime = [System.DateTimeOffset]::Now
$ExportCSV=".\MEM_EnrolledDevices_" + $hours + "h_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

if ($Days -eq $null -or $Days -eq "") {

    #Default is 30 days
    $minutes = 43200 
    $hours = $minutes / 60
    $Days = $hours / 24   

}
else{

    $hours = $Days * 24
    $minutes = $hours * 60 

}

$minutesago = "{0:s}" -f (get-date).addminutes(0-$minutes) + "Z"

write-host "Checking if any Intune Managed Device Enrolled Date is within or equal to $hours hours..." -f Yellow
Write-Host
write-host "Minutes Ago:" $minutesago -f Magenta
Write-Host

try {

    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=enrolledDateTime ge $minutesago"

    $DevicesResponse = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get)

    $Devices = $DevicesResponse.Value
    
    
    #@jmanuelnieto: Dealing with pagination
    $DevicesNextLink = $DevicesResponse."@odata.nextLink"

    while ($null -ne $DevicesNextLink){

            $DevicesResponse = (Invoke-RestMethod -Uri $DevicesNextLink -Headers $authToken -Method Get)
            $DevicesNextLink = $DevicesResponse."@odata.nextLink"
            $Devices += $DevicesResponse.value

        }

    $Devices = $Devices | Sort-Object deviceName
    $Devices = $Devices | Where-Object { $_.managementAgent -ne "eas" }


    # If there are devices not synced in the past 30 days script continues 
    if($Devices){

        $DeviceCount = @($Devices).count    
        Write-Host "There are" $DeviceCount "devices enrolled in the past $hours hours..." -ForegroundColor green
        Write-Host  
        
        
        # Looping through all the devices returned     
        foreach($Device in $Devices){

            $i = 1
            $j = 1
            $activityMsg = "Getting information from Microsoft Endpoint Manager via Microsoft Graph"
            $DeviceID = $Device.id
            #Need to update: To get ChassisType, need to create a function.
            $LSD = $Device.lastSyncDateTime
            $EDT = $Device.enrolledDateTime

            
            $AzDeviceId = $Device.azureADDeviceID
            $AzObjectId = ""
            [string]$deviceAdGroups = ""
            $AzUserId = $Device.userId
            [string]$userAdGroups = ""

            # Painful query 1 by 1, have to specifically select "hardwareInformation" for each device or it comes out "null"
            #  Saving this line, might work in the future: 
            #            $DeviceHardware = $Device.hardwareInformation | Select-Object -Property osBuildNumber,operatingSystemLanguage,operatingSystemEdition,ipAddressV4,subnetAddress,phoneNumber,subscriberCarrier,cellularTechnology,imei

            # Special thanks to Ben Hopper (tw:@BenHopperAU) for helping me solve the $null trauma.

            $DeviceHardware = Get-DeviceHardware $DeviceID
            $DeviceHardware = $DeviceHardware.hardwareInformation | Select-Object -Property osBuildNumber,operatingSystemLanguage,operatingSystemEdition,ipAddressV4,wiredIPv4Addresses,subnetAddress,phoneNumber,subscriberCarrier,cellularTechnology,imei
            
            # Convert collection to string
            $DeviceHardwareWiredIpV4 = $DeviceHardware.wiredIPv4Addresses
            if ($DeviceHardwareWiredIpV4.count -gt 0) {
                
                $ipV4 = ""
                
                foreach ($ip in $DeviceHardwareWiredIpV4) {
                    if ($ipV4 -ne "") { $ipV4 +=", "}
                    $ipV4 += $ip
                
                }

                $DeviceHardwareWiredIpV4 = $ipV4

            }
            else {
                $DeviceHardwareWiredIpV4 = ""
            }

            $EnrolledTime = [datetimeoffset]::Parse($EDT)
            $TimeDifference = $CurrentTime - $EnrolledTime
            $TotalMinutes = ($TimeDifference.TotalMinutes).tostring().split(".")[0]
            
            #Get Device information details, including which groups de device belongs to.
            $DeviceGroups = Get-DeviceGroups $AzDeviceId
            $DeviceGroupsCount = @($DeviceGroups).Count
            $DeviceInfo = Get-DeviceInfo $AzDeviceId
            
            #Get device's primary user information.
            $UserInfo = Get-UserInfo $AzUserId

            #Get AzureAD Group information for device's primary user, but if the user is not found, document user as not found and don't look for groups.  
            if (!($UserInfo -eq "NOT FOUND")) {
                $UserGroups = Get-UserGroups $AzUserId
            }
            else {
             
                #$UserInfo = $UserInfo | Add-Member -NotePropertyMembers @{id = ""; displayName = "N/A"; accountEnabled = "NOT_FOUND"} -Force
                $UserInfo = New-Object psobject -Property @{id = ""; displayName = "N/A"; accountEnabled = "NOT_FOUND"}
                $UserGroups = "USER NOT FOUND"
            }

            $UserGroupsCount = @($UserGroups).Count

            #Share current export status
            $statusMsg = "Currently Processing: " + $AzDeviceID
            Write-Progress -Id 1 -Activity $activityMsg -Status $statusMsg


            #Get name of all Groups that device belongs to, delimit name with simple quotes, put them in a string variable
            foreach ($DeviceGroup in $DeviceGroups) {
                
                $deviceAdGroups += "'"+$DeviceGroup.displayName+"'"
                
                if (!($i -ge $DeviceGroupsCount)) {

                    $deviceAdGroups += " "
                    $i = $i + 1

                }

            }


            #Get name of all Groups that usesr belongs to, delimit name with simple quotes, put them in a string variable
            foreach ($UserGroup in $UserGroups) {
                
                $userAdGroups += "'"+$UserGroup.displayName+"'"
                
                if (!($j -ge $UserGroupsCount)) {

                    $userAdGroups += " "
                    $j = $j + 1

                }

            }

            #Store Device info in table
            $Result = @{'DeviceName'=$Device.deviceName;'AzureDeviceId'=$AzDeviceId;'IntuneDeviceId'=$DeviceID;'DeviceOwnerType'=$Device.managedDeviceOwnerType;'ManagementState'=$Device.managementState;'ManagementAgent'=$Device.managementAgent;'EnrolledProfile'=$Device.enrollmentProfileName;
            'OperatingSystem'=$Device.operatingSystem;'OsSku'=$Device.skuFamily;'DeviceType'=$Device.deviceType;'DeviceChassis'=$Device.chassisType;'LastSyncDateTime'=$Device.lastSyncDateTime;'EnrolledDateTime'=$Device.enrolledDateTime;
            'JailBroken'=$Device.jailbroken;'ComplianceState'=$Device.complianceState;'EnrollmentType'=$Device.deviceEnrollmentType;'AADregistered'=$Device.aadRegistered;'DeviceGroups'=$deviceAdGroups;'DeviceEnabled'=$DeviceInfo.accountEnabled;
            'DeviceDisplayName'=$DeviceInfo.displayName;'DeviceManufacturer'=$DeviceInfo.manufacturer;'DeviceModel'=$Device.model;'DeviceOS'=$DeviceInfo.operatingSystem;'DeviceOSversion'=$DeviceInfo.operatingSystemVersion;'DeviceOSbuild'=$DeviceHardware.osBuildNumber;
            'DeviceOSEdition'=$DeviceHardware.operatingSystemEdition;'DeviceOSlanguage'=$DeviceHardware.operatingSystemLanguage;'DeviceIpV4'=$DeviceHardware.ipAddressV4;'DeviceWiredIpV4'=$DeviceHardwareWiredIpV4;'DeviceSubnet'=$DeviceHardware.subnetAddress;
            'DevicePhoneNumber'=$Device.phoneNumber;'DeviceCarrier'=$DeviceHardware.subscriberCarrier;'DeviceCellTechnology'=$DeviceHardware.cellularTechnology;'AzureUserId'=$AzUserId;'UserGroups'=$userAdGroups;
            'UserEnabled'=$userInfo.accountEnabled;'UserDisplayName'=$userInfo.displayName;'UserCompany'=$userInfo.companyName;'UserCountry'=$userInfo.country;'UserCity'=$userInfo.city;'UserUsageLocation'=$userInfo.usageLocation}

            #Filter for a specific Azure Acite Directory Group. 
            #   Sorry! for sure there should be a better way
            
            if (!($GroupFilter -eq $null -or $GroupFilter -eq "")) { 
                
                if ($deviceAdGroups -match "'$GroupFilter'+" -or $userAdGroups -match "'$GroupFilter'+") {

                    $Results = New-Object PSObject -Property $Result

                    #Informa progress on screen
                    $statusMsg = "Exporting device details: " + $Result.DeviceName
                    Write-Progress -Id 1 -Activity $activityMsg -Status  $statusMsg

                    #Export to file
                    $Results | Select-Object DeviceName,AzureDeviceId,IntuneDeviceId,DeviceOwnerType,ManagementState,ManagementAgent,EnrolledProfile,OperatingSystem,OsSku,DeviceType,DeviceChassis,LastSyncDateTime,EnrolledDateTime,
                    Jailbroken,ComplianceState,EnrollmentType,AADregistered,DeviceGroups,DeviceEnabled,DeviceDisplayName,DeviceManufacturer,DeviceModel,DeviceOS,DeviceOSversion,DeviceOSbuild,DeviceOSEdition,DeviceOSlanguage,
                    DeviceIpV4,DeviceWiredIpV4,DeviceSubnet,DevicePhoneNumber,DeviceCarrier,DeviceCellTechnology,AzureUserId,UserGroups,UserEnabled,UserDisplayName,UserCompany,
                    UserCountry,UserCity,UserUsageLocation | Export-Csv -Path $ExportCSV -Encoding utf8 -Notype -Append

                }
            
            }

            else {

                
                $Results = New-Object PSObject -Property $Result

                #Informa progress on screen
                $statusMsg = "Exporting device details: " + $Result.DeviceName
                Write-Progress -Id 1 -Activity $activityMsg -Status $statusMsg

                #Export to file
                $Results | Select-Object DeviceName,AzureDeviceId,IntuneDeviceId,DeviceOwnerType,ManagementState,ManagementAgent,EnrolledProfile,OperatingSystem,OsSku,DeviceType,DeviceChassis,LastSyncDateTime,EnrolledDateTime,
                Jailbroken,ComplianceState,EnrollmentType,AADregistered,DeviceGroups,DeviceEnabled,DeviceDisplayName,DeviceManufacturer,DeviceModel,DeviceOS,DeviceOSversion,DeviceOSbuild,DeviceOSEdition,DeviceOSlanguage,
                DeviceIpV4,DeviceWiredIpV4,DeviceSubnet,DevicePhoneNumber,DeviceCarrier,DeviceCellTechnology,AzureUserId,UserGroups,UserEnabled,UserDisplayName,UserCompany,
                UserCountry,UserCity,UserUsageLocation | Export-Csv -Path $ExportCSV -Encoding utf8 -Notype -Append

            }


        }

        $msg = "CSV file created in: `n`t" + $PSScriptRoot + $ExportCSV.Substring(1)
        Write-Host $msg -f cyan
        Write-Host

    }

    else {

    write-host "No Devices not checked in the last $minutes minutes found..." -f green
    Write-Host

    }

}

catch {

    Write-Host
    Get-UriCallError $_.Exception $uri "Main section"

}

#Endregion Main
