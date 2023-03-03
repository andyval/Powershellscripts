Param(
    [string]$SoftwareName = "SoftwareNameLookup",
    [Version]$SoftwareVersion = "SoftwareVersionNumer",
    [bool]$ExactSoftwareName = $false,
    [bool]$UpgradeOnly = $false
)

$main = {
    $isInstalled = $false
    $apps = $null
    if($ExactSoftwareName){
	$apps = Get-InstalledSoftware -name $SoftwareName -Exact
    }else{
    	$apps = Get-InstalledSoftware -name $SoftwareName
    }
    foreach($app in $apps){
        if([version]$app.Version -ge [version]$SoftwareVersion){
            $isInstalled = $true
        }
    }
    if($isInstalled){
        Write-Output "Already Installed!"
    }
    if($UpgradeOnly){
        if($null -eq $apps){
            Write-Output "Not Installed, Please Ignore!"
        }
    }
}


#Functions here##
Function Get-InstalledSoftware {
    <#
    .SYNOPSIS
        Retrieves a list of all software installed
    .EXAMPLE
        Get-InstalledSoftware

        This example retrieves all software installed on the local computer
    .PARAMETER Name
        The software title you'd like to limit the query to.
    #>
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
	[switch]$exact
    )

    $UninstallKeys = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    $null = New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS
    $UninstallKeys += Get-ChildItem HKU: -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object { "HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall" }
    if (-not $UninstallKeys) {
        Write-Verbose -Message 'No software registry keys found'
    } else {
        foreach ($UninstallKey in $UninstallKeys) {
            if ($PSBoundParameters.ContainsKey('Name')) {
                $WhereBlock = { ($_.PSChildName -like '*') -and ($_.GetValue('DisplayName') -like "$Name*") }
		if($exact){
		 	 $WhereBlock = { ($_.PSChildName -like '*') -and ($_.GetValue('DisplayName') -eq "$Name") }
		}
            } else {
                $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName')) }
            }
            $gciParams = @{
                Path        = $UninstallKey
                ErrorAction = 'SilentlyContinue'
            }
            $selectProperties = @(
                @{n='GUID'; e={$_.PSChildName}}, 
                @{n='Name'; e={$_.GetValue('DisplayName')}},
                @{n='Version'; e={$_.GetValue('DisplayVersion')}},
                @{n='UninstallString'; e={$_.GetValue('UninstallString')}}
            )
            Get-ChildItem @gciParams | Where $WhereBlock | Select-Object -Property $selectProperties
        }
    }
}

#Invoke Main
. $main
