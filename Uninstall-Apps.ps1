<#
.Synopsis
   Uninstall Applications based on the product code/GUID or the FileVersion

#>


#---------------------INITIALIZATION----------------------------------#
Set-ExecutionPolicy bypass -Scope process -Force
#---------------------vARIABLES---------------------------------------#
$appName = "Displaylink"
$detectionRuleType = "FilePath" ##ProductCode or FilePath
$ErrorActionPreference = 'Continue'

$ProductCode = ""
$appFilePathDetectionRule = "C:\Program Files\DisplayLink Core Software\DisplayLinkTrayApp.exe" #Check to see if Latest Version is installed
$FileVersion = "9.0.1651.0"
#---------------------FUNCTIONS---------------------------------------#
function Get-InstalledSoftware {
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
        [string]$Name
    )
 
    $UninstallKeys = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    $null = New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS
    $UninstallKeys += Get-ChildItem HKU: -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object { "HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall" }
    if (-not $UninstallKeys) {
        Write-Verbose -Message 'No software registry keys found'
    } else {
        foreach ($UninstallKey in $UninstallKeys) {
            if ($PSBoundParameters.ContainsKey('Name')) {
                $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName') -like "$Name*") }
            } else {
                $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName')) }
            }
            $gciParams = @{
                Path        = $UninstallKey
                ErrorAction = 'SilentlyContinue'
            }
            $selectProperties = @(
                @{n='GUID'; e={$_.PSChildName}}, 
                @{n='Name'; e={$_.GetValue('DisplayName')}}
                @{n='Version'; e={$_.GetValue('DisplayVersion')}}
            )
            Get-ChildItem @gciParams | Where $WhereBlock | Select-Object -Property $selectProperties
        }
    }
}

function Test-FileVersion{
    <#
    .SYNOPSIS
       Tests to see if the version of a file is greater than, less than, or equal to 
       the version number defined. 
    .EXAMPLE
        Test-FileVersion -filename "C:\Windows\explorer.exe" -version "1.0.0.0"
        
        This example retrieves all software installed on the local computer
    .PARAMETER filename
        The path of the file
    .PARAMETER version
        The version number you want to query against the file. 
    #>
    Param(
          $filename=$null,
          $version,
          $productcodeVersion=$null
    )
    $version = [System.Version]::Parse($version)

    if($filename -eq $null -and $productcodeVersion -eq $null){
        return Write-Error "Please enter filename or productcode version" 
    }
    
    if($filename -ne $null){
    $versionInfo = (Get-Item "$($filename)").VersionInfo
    $fileVersion = ("{0}.{1}.{2}.{3}" -f $versionInfo.FileMajorPart, $versionInfo.FileMinorPart, $versionInfo.FileBuildPart, $versionInfo.FilePrivatePart)
    $fileVersion = [System.Version]::Parse($FileVersion)
    }

    if($productcodeVersion -ne $null){
        $fileVersion = [System.Version]::Parse($productcodeVersion)
    }
    
    if($fileVersion -gt $version){
        return "greater"
    }
    if($fileVersion -lt $version){
        return "less than"
    }
    if($fileVersion -eq $version){
        return "equal"
    }
    else {
        return "error"
    }
}

#--------------------------Main--------------------------------------#
If($appName -eq ""){
    Write-Host "No App Name Given." -ForegroundColor Red
    exit 
}

If($detectionRuleType -ne "FilePath" -or $detectionRuleType -ne "ProductCode"){
    Write-Host "Please choose FilePath or ProductCode for detection rule type" -ForegroundColor Red
    exit
}
If($ProductCode -eq "" -and $appFilePathDetectionRule -eq ""){
    Write-Host "You must enter a FilePath or ProductCode" -ForegroundColor Red
    exit
}

switch($detectionRuleType){
    'FilePath' {
        If($appFilePathDetectionRule -eq ""){
            Write-Host "No File Path Given" -ForegroundColor Red
            exit
        }
        else {
            If($FileVersion -eq ""){
                Write-Host "No file Version given" -ForegroundColor Red
                exit
            }

            $versionCheck = Test-FileVersion -filename $appFilePathDetectionRule -version $FileVersion
            switch($versionCheck){
                'greater' {Write-Host "A more recent version is installed!" -ForegroundColor Green;break}
                'equal' {Write-Host "Product already installed!" -ForegroundColor Green;break}
                'error' {Write-Host "Error, please check version and file path" -ForegroundColor red;break}
                'less than'{
                        $apps = Get-InstalledSoftware -Name $appName
                        If($apps -eq $null){
                            Write-Host "No apps found." 
                            break;
                        }
                        else {
                            Write-Host "Uninstalling Older Versions"
                            foreach ($app in $apps){
                                (Start-Process -file "msiexec.exe" -ArgumentList " /X $($app.GUID) /qn /norestart" -wait -passthru).ExitCode
                            }
                        }
                    }
                 
            }
            
        }
        
    }
    'ProductCode' {
        If ($ProductCode -eq ""){
        Write-Host "No product code entered"
        exit
        }
        else {
            $apps = Get-InstalledSoftware -Name $appName
            #check to make sure Get-Instaled Software returns a value
            if ($apps -eq $null){
                Write-Host "Application not installed"
                exit
            }
            #checks to see if application has already been installed
            foreach($app in $apps){
                if ($app.GUID -eq $ProductCode){
                    Write-Host "Software Already installed"
                    exit
                }
                #if FileVersion was defined
                if($FileVersion -ne ""){
                    $versionCheck = Test-FileVersion -version $FileVersion -productcodeVersion $app.Version
                    #check to see if later version was installed
                    if($versionCheck -eq 'greater'){
                        Write-Host "A more recent version is installed!"
                        exit
                    }
                }
            }
            #Loop thru to uninstall ALL applications matching $appName
            foreach($app in $apps){
                    (Start-Process -file "msiexec.exe" -ArgumentList " /X $($app.GUID) /qn /norestart" -wait -passthru).ExitCode
            }
        }
    
    }
}
