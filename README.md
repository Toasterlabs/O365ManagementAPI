# O365ManagementAPI

Collection of PowerShell code to retrieve information from Office 365 using the O365 management API...

To use, you'll need to setup an app in O365 andconfigure the configuration file:

{
    "Description": "Global configuration file for O365 Investigations",
    "AppID": "<App ID goes here>",
    "AppSecret": "<App key goes here>",
    "TenantDomain": "<your tenant domain (contoso.onmicrosoft.com or contoso.com)>",
    "TenantGUID": "<This is not a app identifier, get this from azure portal - Azure Active Directory - Properties - Directory ID >",
    "LoginURL" : "https://login.windows.net",
    "ResourceAPI" : "https://manage.office.com"
}
