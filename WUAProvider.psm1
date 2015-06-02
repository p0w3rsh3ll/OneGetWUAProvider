$Providername = "WUA"

Function Initialize-Provider     { Write-Debug "In $($Providername) - Initialize-Provider" }
Function Get-PackageProviderName { Return $Providername }

Function Resolve-PackageSource {
Param()

    $dbgh = "In $($Providername) - Resolve-PackageSources"
    Write-Debug $dbgh 

    $IsTrusted    = $true
    $IsRegistered = $true
    $IsValidated  = $true

    $UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
        
    $UpdateSvc.Services | 
    Where { $_.Name -eq 'Windows Update'} | 
    ForEach-Object {
        Write-Debug "$dbgh - Resolved WUA API source $($_.Name) to: $($_.ServiceID)" 
        New-PackageSource $($_.Name) $($_.ServiceID) $IsTrusted $IsRegistered $IsValidated
    }
}

Function Find-Package {
    Param(
        [string[]] $names,
        [string] $requiredVersion,
        [string] $minimumVersion,
        [string] $maximumVersion
    )

	$dbgh = "In $($Providername) - Find-Package"
    Write-Debug $dbgh

    $svcid = Resolve-PackageSource
    Write-Debug "$dbgh - Using resolved source $($svcid.Name)" 

    try {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        $Searcher.ServiceID = $svcid.Location ; # location = ServiceId

        $Criteria = "IsInstalled=0 and DeploymentAction='Installation' or 
        IsPresent=1 and DeploymentAction='Uninstallation' or
        IsInstalled=1 and DeploymentAction='Installation' and RebootRequired=1 or 
        IsInstalled=0 and DeploymentAction='Uninstallation' and RebootRequired=1"
            
        $SearchResult = $Searcher.Search($Criteria)
            
        if ($SearchResult.ResultCode -eq 2) {
            Write-Debug "$dbgh - Found $(@($SearchResult.Updates).count) updates from source: $($svcid.Name)"

            @($SearchResult.Updates) | ForEach-Object {
                $u = $_ ;

                Write-Debug "$dbgh - Found update from source $($svcid.Name): $($u.Title)"
                $names | ForEach-Object {
                    if ($u.Title -like "*$($_)*") {
                        $swid = @{
                            fastPackageReference = $($u.Identity.UpdateID) ;
                            version = "1.0.0.$($u.Identity.RevisionNumber)" ;
                            versionScheme = 'semver' ;
                            name = $u.Title ;
                            source = $($svcid.Name) ;
                            summary = $u.Description ;
                            fromTrustedSource  = $true ;
                            details = @{ 
                                UpdateObj = $u ;
                                Id = $($u.Identity.UpdateID) ;
                            } ; # to make it installable w/o searching
                        }
                        New-SoftwareIdentity @swid
                    }
                }
            }
        }
    } catch {
        throw $_
    }    

} 

Function Get-InstalledPackage {
Param()
    $dbgh = "In $($Providername) - Get-InstalledPackage"
    Write-Debug $dbgh
    try {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        $HistoryCount = $Searcher.GetTotalHistoryCount()
        Write-Debug "$dbgh - About to list installed packages from $HistoryCount total WUA based packages"
        $Searcher.QueryHistory(0,$HistoryCount) | ForEach-Object -Process {
            # IUpdateHistoryEntry > https://msdn.microsoft.com/en-us/library/windows/desktop/aa386400(v=vs.85).aspx
            if ($_.ResultCode -eq 2) {
                $swid = @{
                    FastPackageReference = $_.UpdateIdentity.UpdateID ;
                    Name = $_.Title ;
                    Version = "1.0.0.$($_.UpdateIdentity.RevisionNumber)" ;
                    VersionScheme = 'semver' ;
                    Source = 'local';
                    Summary = $_.Description ;
                    FromTrustedSource  = $true ;
                    Details = @{ 
                        UpdateObj = $_ ;
                        Id = $_.UpdateIdentity.UpdateID ;
                    } ; # to make it installable w/o searching
                }
                New-SoftwareIdentity @swid
            }
        } 
    } catch {
        throw $_
    }
}

Function Install-Package {
Param(
    [string] $fastPackageReference
)

    $dbgh = "In $($Providername) - Install-Package"
    Write-Debug $dbgh

    $Session = New-Object -ComObject Microsoft.Update.Session
    $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    $updatesToDownload.clear()
    $downloader = $Session.CreateUpdateDownloader()
    $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $updatesToInstall.clear()

    if ($fastPackageReference) {
        Write-Debug "$dbgh - About to find packages using fastPackageReference: $fastPackageReference"
        $pkg = Find-Package -Source (Resolve-PackageSource) | Where fastPackageReference -eq $fastPackageReference
        Write-Debug "$dbgh - Found $($pkg.Count) matching $fastPackageReference"

        $pkg | ForEach-Object {
            $null = $updatesToDownload.Add(
                $_.Details['UpdateObj']
            ) ; # Only IUpdate can be added
        }
        if ($updatesToDownload.Count -ge 1) {
            $downloader.Updates = $updatesToDownload
            Write-Debug "$dbgh - Downloading update $fastPackageReference" 
            $downresults = $downloader.Download()

            if ($downresults.ResultCode -eq 2) {

                Write-Debug "$dbgh - Successfully downloaded update $fastPackageReference" 
                $updatesToDownload | 
                Where { -not($_.isInstalled) -and $_.isDownloaded} | ForEach-Object {
                    $updatesToInstall.Add($_) | Out-Null
                }
                $installer = $Session.CreateUpdateInstaller()
                $installer.Updates = $updatesToInstall
                Write-Debug "$dbgh - Installing update $fastPackageReference" 
                $installationResult = $installer.Install()
            
                if ($installationResult.ResultCode -eq 2) {
                    $pkg
                }
            } else {
                Write-Debug "$dbgh - Downloading update $fastPackageReference failed with result code $($installationResult.ResultCode)" 
            }
        }
    }
}
