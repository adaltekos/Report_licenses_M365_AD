#Install-Module SharePointPnPPowerShellOnline
#Install-Module Microsoft.Graph
#Install-Module ImportExcel
#Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online

$filename 		= "" #Complete with filename (ex. Raport_M365_users_services.xlsx)
$localPath 		= "" #Complete with local path (ex. C:\Raporty\)
$siteUrl		= "" #Complete with Url site (ex. https://company.sharepoint.com/sites/it-dep)
$onlinePath		= "" #Complete with path where file is on sharepoint (ex. Shared Documents/Global/)
$tenant			= "" #Complete with tenant name (ex. company.onmicrosoft.com)
$appId			= "" #Complete with ClientId (which is ID of application registered in Azure AD)
$thumbprint		= "" #Complete with Thumbprint (which is certificate thumbprint)

$pnpConnectParams  = @{
    Url				=  $siteUrl
    Tenant			=  $tenant
    ClientId		=  $appId
    Thumbprint		=  $thumbprint
}
Connect-PnPOnline @pnpConnectParams

$getPnPFileParams = @{
    Url				= ($onlinePath + $filename)
    Path			= $localPath
    Filename		= $filename
    AsFile			= $true
    Force			= $true
}
Get-PnPFile @getPnPFileParams

Start-Sleep -s 3

$graphParams  = @{
    Tenant					= $tenant
    AppId					= $appId
    CertificateThumbprint	= $thumbprint
}
Connect-Graph @graphParams

Get-MgSubscribedSku | Select -Property SkuPartNumber, ConsumedUnits, @{Name='Enabled'; Expression={$_.PrepaidUnits.Enabled}}, @{Name='Suspended'; Expression={$_.PrepaidUnits.Suspended}}, @{Name='Warning'; Expression={$_.PrepaidUnits.Warning}} | Where-Object {($_.SkuPartNumber -eq "O365_BUSINESS_ESSENTIALS") -or ($_.SkuPartNumber -eq "O365_BUSINESS_PREMIUM")} | Export-Excel -Path ($onlinePath + $filename) -WorkSheetname new -AutoSize

Start-Sleep -s 3

#Excel
$excel = Open-ExcelPackage -Path ($onlinePath + $filename)

#M365
$excel.old.Cells["A4:E5"].Value = $excel.old.Cells["A2:E3"].Value
$excel.old.Cells["C2:D3"].Value = $excel.new.Cells["B2:C3"].Value
$excel.old.Cells["A2"].Value = $(Get-Date -Format "MMM.yy")

#AD
$excel.old.Cells["G3:K3"].Value = $excel.old.Cells["G2:K2"].Value
$excel.old.Cells["I2"].Value = $((Get-ADUser -Filter *).count)
$excel.old.Cells["G2"].Value = $(Get-Date -Format "MMM.yy")

#Historia
$excel.Historia.Cells["G1:BI5"].Value = $excel.Historia.Cells["B1:BD5"].Value
$excel.Historia.Cells["B1"].Value = $(Get-Date -Format "MMM.yy")
$excel.Historia.Cells["D3:D4"].Value = $excel.old.Cells["C2:C3"].Value
$excel.Historia.Cells["B3:B4"].Value = $excel.old.Cells["D2:D3"].Value
$excel.Historia.Cells["F3:F4"].Value = $excel.old.Cells["E2:E3"].Value
$excel.Historia.Cells["D5"].Value = $excel.old.Cells["I2"].Value
$excel.Historia.Cells["B5"].Value = $excel.old.Cells["J2"].Value
$excel.Historia.Cells["F5"].Value = $excel.old.Cells["K2"].Value
$excel.Historia.Cells["C3"].Value = $excel.Historia.Cells["B3"].Value - $excel.Historia.Cells["G3"].Value
$excel.Historia.Cells["C4"].Value = $excel.Historia.Cells["B4"].Value - $excel.Historia.Cells["G4"].Value
$excel.Historia.Cells["C5"].Value = $excel.Historia.Cells["B5"].Value - $excel.Historia.Cells["G5"].Value
$excel.Historia.Cells["E3"].Value = $excel.Historia.Cells["D3"].Value - $excel.Historia.Cells["I3"].Value
$excel.Historia.Cells["E4"].Value = $excel.Historia.Cells["D4"].Value - $excel.Historia.Cells["I4"].Value
$excel.Historia.Cells["E5"].Value = $excel.Historia.Cells["D5"].Value - $excel.Historia.Cells["I5"].Value

Close-ExcelPackage -ExcelPackage $excel

$addPnPFileParams  = @{
    Folder		= $onlinePath	
    Path		= ($localPath + $filename)
}
Add-PnPFile @addPnPFileParams