<#
.SYNOPSIS
    Script to create WMI Filters via powershell

.DESCRIPTION
    This script creates a WMI Filter in Active Directory using the LDIFDE utility. 
    It generates a unique GUID for the filter and sets the necessary attributes based on the provided parameters.
    By default, the script outputs the LDIF content to the console. If the -import switch is used, it imports the filter directly into Active Directory.

.NOTES
    To be able to import the WMI Filter you need to have the appropriate permissions in Active Directory and ensure that the LDIFDE utility is available on your system.

.PARAMETER name
    The name of the WMI Filter to be created.

.PARAMETER description
    A description for the WMI Filter.

.PARAMETER query
    The WMI query that defines the filter criteria.

.PARAMETER author
    The author of the WMI Filter. Default is "Administrator".

.PARAMETER import
    If specified, the script will import the WMI Filter into Active Directory instead of outputting the LDIF content to the console.    

.LINK
    https://github.com/hethdev/AD-Tools
    
.EXAMPLE
    .\CreateWMIFilter.ps1 -name "Windows 10 Filter" -description "Filter for Windows 10 devices" -query "SELECT * FROM Win32_OperatingSystem WHERE Version LIKE '10.%' AND ProductType='1'" -import
    This command creates a WMI Filter named "Windows 10 Filter" with the specified description and query, and imports it into Active Directory.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)][string]$name,
    [Parameter(Mandatory)][string]$description,
    [Parameter(Mandatory)][string]$query,
    [string]$author = "Administrator",
    [switch]$import
)

if(-not (test-path "C:\windows\System32\ldifde.exe")) {
    Write-Error "LDIFDE utility not found. Please ensure it is available on your system."
    exit
}

$guid = (New-Guid).Guid
$date = Get-Date -Format "yyyyMMddHHmmss.281000-000"
$domainDN = (Get-ADDomain).DistinguishedName

$object = @"
dn: CN={$guid},CN=SOM,CN=WMIPolicy,CN=System,$domainDN
objectClass: msWMI-Som
cn: {$guid}
distinguishedName: CN={$guid},CN=SOM,CN=WMIPolicy,CN=System,$domainDN
name: {$guid}
objectCategory: CN=ms-WMI-Som,CN=Schema,CN=Configuration,$domainDN
msWMI-Author: $author
msWMI-ChangeDate: $date
msWMI-CreationDate: $date
msWMI-ID: {$guid}
msWMI-Name: $name
msWMI-Parm1: $description
msWMI-Parm2: 1;3;10;18;WQL;root\CIMv2;$query;

"@

if ($import) { 
    $object | Out-file ($output = New-TemporaryFile)
    $ldifde = LDIFDE -i -f $($output.FullName)
    $ldifde | ForEach-Object { Write-Verbose $_ }
    if ($LASTEXITCODE -eq '0') {
        Write-Host "WMI Filter Successfully Imported" -ForegroundColor Green
    } else {
        Write-Host "WMI Filter Import Failed" -ForegroundColor Red
    }
}
else {
    return $object
}