# Report_licenses_M365_AD

## Description

This PowerShell script retrieves the number of subscribed SKU information (licenses) from Microsoft Graph and the number of users from AD to compare with the purchased AD licenses. It exports the information to an Excel file stored in SharePoint Online.

## Prerequisites
- PowerShell version 5.1 or later
- Installed SharePointPnPPowerShellOnline module
- Installed Microsoft.Graph module
- Installed ImportExcel module
- Installed Remote Server Administration Tools (RSAT) for Active Directory

## Configuration

The script uses the following variables which need to be set:

- `$filename`: The name of the Excel file to be generated (ex. Raport_M365_users_services.xlsx)
- `$localPath`: The local path where the Excel file will be saved (ex. C:\Raporty\)
- `$siteUrl`: The URL of the SharePoint Online site where the Excel file will be stored (ex. https://company.sharepoint.com/sites/it-dep)
- `$onlinePath`: The path where the Excel file will be stored on SharePoint Online (ex. Shared Documents/Global/)
- `$tenant`: The name of the tenant (ex. company.onmicrosoft.com)
- `$appId`: The Client ID of the application registered in Azure AD
- `$thumbprint`: The certificate thumbprint
