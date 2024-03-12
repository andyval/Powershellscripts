Param (
    [string]$LogFolder = "C:\Apps\Logs",
    [ValidateSet("x86","x64")]
    [string]$arch = "x64",
    [ValidateSet("True","False")]
    [string]$MigrateArch = "True",
    [string] $downloadfolder = "C:\apps\M365",
    [string]$setupEXEurl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
)
$main = {
    [System.Net.WebClient]$webClient = New-Object System.Net.WebClient
    if(!(Test-path $LogFolder)){new-item -ItemType Directory -Path $LogFolder -Force}
    if(!(Test-path $downloadfolder)){new-item -ItemType Directory -Path $downloadfolder -Force}
    #Check for Channel installed
    #First Check GPO, then check installed value
    $GPOchannel = get-itemproperty "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" -Name "updatebranch" -ErrorAction SilentlyContinue
    $InstalledChannel = get-itemproperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "AudienceId" -ErrorAction SilentlyContinue
    if($null -ne $GPOchannel){
        $channel_lookup = @{
            'MonthlyEnterprise'     = 'MonthlyEnterprise'
            'Current'               = 'Current'
            'FirstReleaseCurrent'   = 'CurrentPreview'
            'Deferred'              = 'SemiAnnual'
            'FirstReleaseDeferred'  = 'SemiAnnualPreview' 
            'InsiderFast'           = 'BetaChannel'
        }
        $channel = ($GPOchannel | select @{n= 'Channel';e = {$channel_lookup.item($_.updatebranch)}}).Channel

    }elseif($null -ne $InstalledChannel){
        #If no GPO has been applied, then lets grab the installed Version Channel name (only applies to PCs with previous installs, not imaging)
        $channel_lookup = @{
            '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'MonthlyEnterprise'
            '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'
            '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'CurrentPreview'
            '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'SemiAnnual'
            'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'SemiAnnualPreview'
            '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'BetaChannel'
        }
        $channel = ($InstalledChannel | select @{n= 'Channel';e = {$channel_lookup.item($_.AudienceId)}}).Channel
    }else{
        #Default to MonthlyEnterprise Channel
        $channel = 'MonthlyEnterprise'
    }
       
    if($null -eq $arch){
        $Architecture = get-itemproperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "Platform" -ErrorAction SilentlyContinue
        if($null -ne $Architecture){
            $arch = $Architecture.Platform
        }else{
            $arch = "x86"
         }
    }
    if($arch -eq "x64"){
        $OfficeClientEdition = "64"
    }else{
        $OfficeClientEdition = "32"
    }
    $XML =  @"
<Configuration>
    <Add OfficeClientEdition="$OfficeClientEdition" Channel="$channel" AllowCdnFallback="True" MigrateArch="$MigrateArch">
        <Product ID="O365ProPlusRetail">
            <Language ID="MatchOS" Fallback="en-us" />
            <ExcludeApp ID="Groove" />
            <ExcludeApp ID="Lync" />
            <ExcludeApp ID="Bing" />
            <ExcludeApp ID="OneDrive" />
        </Product>
        <Product ID="VisioProRetail">
            <Language ID="MatchOS" Fallback="en-us" />
        </Product>
        <Product ID="ProjectProRetail">
            <Language ID="MatchOS" Fallback="en-us" />
        </Product>
    </Add>
    <Property Name="SharedComputerLicensing" Value="0" />
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Updates Enabled="TRUE" />
    <RemoveMSI />
    <Display Level="Full" AcceptEULA="TRUE" />
    <AppSettings>
        <User Key="software\microsoft\office\16.0\outlook\options\general" Name="DisablePreviewPlace" Value="1" Type="REG_DWORD" App="outlk16" Id="L_DisablePreviewPlace" />
    </AppSettings>
</Configuration>
"@
    [System.IO.File]::WriteAllLines("$downloadfolder\O365.XML", $XML, $(New-Object System.Text.UTF8Encoding $False))
    

    #downloads the latest setup.exe
    $webClient.DownloadFile($setupEXEurl, "$downloadfolder\setup.exe")
    
    #Copy setup.exe from package if 
    if(!(Test-Path "$downloadfolder\setup.exe")){
        Copy-Item -Path "$PSScriptRoot\setup.exe" -Destination "$downloadfolder\setup.exe" -Force -ErrorAction SilentlyContinue
    }
    #if it still doesnt exist then lets throw an error up
    if(!(Test-Path "$downloadfolder\setup.exe")){
        Write-Output "Error downloading setup.exe"
        Exit 1
    }
    #regkey for viewer mode for :
    New-Item "hklm:\software\policies\microsoft\office\16.0\common\licensing" -Force
    New-ItemProperty "hklm:\software\policies\microsoft\office\16.0\common\licensing" -Name "viewermode" -PropertyType DWORD -Value 1 -Force

    #install
    Set-location $downloadfolder
    Start-Process "$downloadfolder\setup.exe" -ArgumentList "/configure O365.XML" -wait -PassThru -NoNewWindow
    if($MigrateArch -eq "True"){
        $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
        do{
            $Architecture = get-itemproperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "Platform" -ErrorAction SilentlyContinue

        }until(($Architecture.Platform -eq $arch) -or ($stopwatch.Elapsed.TotalSeconds -gt 300))
        $stopwatch.stop()
    }
}

. $main
