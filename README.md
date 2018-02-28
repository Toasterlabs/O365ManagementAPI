# O365ManagementAPI

## Description
Collection of PowerShell code to retrieve information from Office 365 using the O365 management API... If you decide to reuse this code in your own projects, give me a shout-out!

Currently has the following functions:

### get-globalconfig
*Param*: ConfigFile
Imports configuration from json

### Get-oauthtoken
*Param*: Resource, ClientID, ClientSecret, LoginURL, TenantDomain
Constructs header for use with other functions

### get-SVcMessages
*Param*: ConfigFile, OutputPath
Pulls all messages down from the Service Health Dashboard and stores them in XML format by ID. Overwrites any existing XML file with the same name

### get-CurrentStatus
*Param*: ConfigFile, Outputpath
Pulls the current status of the service and stores it in a CSV file. This will append any existing CSV that exists!

### get-Services
*Param*: ConfigFile, Outputpath
Pulls the currently enabled services from O365 in to a CSV file. This will append any existing CSV that exists!

##  Configuration
You'll need to set up an app in the Azure management portal: https://msdn.microsoft.com/en-us/office-365/get-started-with-office-365-management-apis?f=255&MSPPError=-2147217396#application-registration-in-azure-ad

```
{
    "Description": "Global configuration file for O365 Investigations",
    "AppID": "<App ID goes here>",
    "AppSecret": "<App key goes here>",
    "TenantDomain": "<your tenant domain (contoso.onmicrosoft.com or contoso.com)>",
    "TenantGUID": "<This is not a app identifier, get this from azure portal - Azure Active Directory - Properties - Directory ID >",
    "LoginURL" : "https://login.windows.net",
    "ResourceAPI" : "https://manage.office.com"
}
```

## Disclaimer
Use this code at your own risk, verify anything you run in production. If you decide to run this code all risk stays with you. 

### Legalese:
> The sample scripts are provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall I,or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if I have been advised of the possibility of such damages.
