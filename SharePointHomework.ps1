###########################################################################################################################################
#                                                                                                                                         #
# SharePoint homework                                                                                                                     #
#                                                                                                                                         #
# To complete the tasks, I've used PnP.PowerShell, because at this moment it is the most modern way to work with SharePoint               #
#                                                                                                                                         #
# Note #1                                                                                                                                 #
# In case of using CSOM, I would first check, if there are necessary assemblies available:                                                #
# $assemblies= @("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll",         #
#                "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll") #
#                                                                                                                                         #
# If not, I would download them:                                                                                                          #
# Invoke-WebRequest -Uri $uri -OutFile $path                                                                                              #
#                                                                                                                                         #
# Note #2                                                                                                                                 #
# When I write such PowerShell scripts, I'm adding checks for existance of the site, library, field etc. before creation or updating.     #
# In such a way it is possible to run the script as many times, as needed - it is useful, if something went wrong and it is necessary     #
# to re-run the script to achive all the needs.                                                                                           #
#                                                                                                                                         #
###########################################################################################################################################

# Specify your SharePoint Online Admin center URL here
$siteUrl = "https://yallamsdn-admin.sharepoint.com"
# and Super Admin user for both sites
$superAdmin = "testuser1@yallamsdn.onmicrosoft.com"
# and Admin user for 'Hub documents' only
$admin = "testuser2@yallamsdn.onmicrosoft.com"

function CreateHubHome() {
    $hubHomeTitle = "Hub Home"
    Write-Host "Creating '$hubHomeTitle'..." -f Green
    $hubHomeUrl = New-PnPSite -Type TeamSite -Title $hubHomeTitle -Alias "hubhome" -Lcid 1033 -Wait
    Write-Host $hubHomeUrl

    Connect-PnPOnline -Url $hubHomeUrl -Interactive  # In my case it didn't ask for credentials, as they alredy provided earlier

    # NOTE: I'm not sure about the lines below, probably something different has been meant by:
    #       "Sets up a super admin user for both sites and another admin user for "Hub documents" site only."
    Add-PnPSiteCollectionAdmin -Owners $superAdmin
    Write-Host "Site Collection Administrators added: $superAdmin"

    Write-Host "Registering the site as a hub site... " -NoNewline
    Register-PnPHubSite -Site $hubHomeUrl
    Write-Host "done" -f Green

    Write-Host "Enabling external sharing... " -NoNewline
    Set-PnPTenantSite -Url $hubHomeUrl -SharingCapability ExternalUserSharingOnly
    Write-Host "done" -f Green

    Write-Host "Applying 'If Design' to the '$hubHomeTitle'... " -NoNewline
    # NOTE: Here I assume, that there is an "If Design" available in the system.
    #       In real project I would write a function, where I would try to get the site design, and if it would not exist, I would create it, e.g. EnsureIfDesign()
    Invoke-PnPSiteDesign -WebUrl $hubHomeUrl -Identity "If Design"
    Write-Host "done" -f Green

    return $hubHomeUrl
}

function CreateHubDocs($hubHomeUrl) {
    $hubDocsTitle = "Hub documents"
    Write-Host "Creating '$($hubDocsTitle)'..." -f Green
    $hubDocsUrl = New-PnPSite -Type TeamSite -Title $hubDocsTitle -Alias "hubdocuments" -Lcid 1033 -Wait
    Write-Host $hubDocsUrl

    Connect-PnPOnline -Url $hubDocsUrl -Interactive  # In my case it didn't ask for credentials, as they alredy provided earlier

    Add-PnPSiteCollectionAdmin -Owners @($admin, $superAdmin)
    Write-Host "Site Collection Administrators added: $(@($admin, $superAdmin))"

    Write-Host "Associating the site with '$hubHomeTitle'... " -NoNewline
    Add-PnPHubSiteAssociation -Site $hubDocsUrl -HubSite $hubHomeUrl
    Write-Host "done" -f Green

    $classifiedLibTitle = "Classified"
    Write-Host "Creating library '$($classifiedLibTitle)'... " -NoNewline
    New-PnPList -Url "classified" -Title $libTitle -Template DocumentLibrary -OnQuickLaunch > $null
    Write-Host "done" -f Green

    $classifiedLib = Get-PnPList -Identity $classifiedLibTitle

    $commentsFieldTitle = "Comments"
    Write-Host "Adding field '$commentsFieldTitle'... " -NoNewline
    Add-PnPField -List $classifiedLib -InternalName $commentsFieldTitle -DisplayName $commentsFieldTitle -Type Text -Required -AddToDefaultView > $null
    Write-Host "done" -f Green

    # NOTE: Unfortunately, I was not able to fully test 'Set-PnPListInformationRightsManagement', as I don't have possibility to activate IRM in my tenant.
    #       I didn't find possibility to activate this requirement through PnP.PowerShell: "Do not allow users to upload documents that do not support IRM".
    #       In such case in real project I would try to use different approach.
    Write-Host "Setting IRM settings for the library '$classifiedLibTitle'... " -NoNewline
    Set-PnPListInformationRightsManagement -List $classifiedLib -PolicyTitle "Classified documents" -Enable $true -AllowPrint $true -EnableRejection $true `
        -EnableDocumentAccessExpire $true -DocumentAccessExpireDays 90 -EnableLicenseCacheExpire $true -LicenseCacheExpireDays 90 > $null
    Write-Host "done" -f Green
}

try {
    # Install the necessary module PnP.PowerShell
    # NOTE: Strange, but Get-Module doesn't return 'PnP.PowerShell' module unless Connect-PnPOnline is called...
    if (Get-Module -Name PnP.PowerShell) {
        Write-Host "PnP.PowerShell already installed" -f Gray
    } else {
        Write-Host "Installing PnP.PowerShell module..."
        Install-Module -Name PnP.PowerShell
    }

    # Connect to the admin center site, using your admin account
    Connect-PnPOnline -Url $siteUrl -Interactive

    $hubHomeUrl = CreateHubHome
    
    CreateHubDocs -hubHomeUrl $hubHomeUrl

} catch [SystemException] {
    Write-Error $_
}