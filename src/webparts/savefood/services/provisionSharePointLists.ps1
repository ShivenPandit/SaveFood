<#
.SYNOPSIS
    Automated SharePoint Lists Provisioning Script for Vedanta Save Food Management System.
.DESCRIPTION
    Creates the 3 required SharePoint Lists on the target site (SaveFoodUAT or SaveFood):
    1. LocationMaster (Single line text Location, Boolean IsActive, seeded with 6 locations)
    2. RoleMatrix (Single line text Title, Person Multi-value Admin)
    3. MealOptOutRequest (Title, Lookup Location, Choice MealType, DateOnly MealDate, Person Requester)
.PARAMETER SiteUrl
    The target SharePoint Site URL (e.g. https://vedanta.sharepoint.com/sites/SaveFoodUAT)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl = "https://vedanta.sharepoint.com/sites/SaveFoodUAT"
)

# Connect to target SharePoint site using PnP.PowerShell
Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# -------------------------------------------------------------------
# 1. LocationMaster List
# -------------------------------------------------------------------
$locationListName = "LocationMaster"
Write-Host "Checking $locationListName..." -ForegroundColor Yellow

$locationList = Get-PnPList -Identity $locationListName -ErrorAction SilentlyContinue
if (-not $locationList) {
    Write-Host "Creating list $locationListName..." -ForegroundColor Green
    $locationList = New-PnPList -Title $locationListName -Template GenericList -EnableVersioning
    Set-PnPList -Identity $locationListName -FieldTitle "Location"
}

# Add IsActive column
$isActiveField = Get-PnPField -List $locationListName -Identity "IsActive" -ErrorAction SilentlyContinue
if (-not $isActiveField) {
    Write-Host "Adding IsActive column to $locationListName..." -ForegroundColor Green
    Add-PnPField -List $locationListName -Type Boolean -DisplayName "Is Active" -InternalName "IsActive" -AddToDefaultView
}

# Seed initial locations
$initialLocations = @("LQ", "MCR", "Project Office", "CCR", "CPF", "Aish")
foreach ($loc in $initialLocations) {
    $existing = Get-PnPListItem -List $locationListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$loc</Value></Eq></Where></Query></View>"
    if (-not $existing) {
        Write-Host "Seeding location: $loc" -ForegroundColor Cyan
        Add-PnPListItem -List $locationListName -Values @{ "Title" = $loc; "IsActive" = $true } | Out-Null
    }
}

# -------------------------------------------------------------------
# 2. RoleMatrix List
# -------------------------------------------------------------------
$roleListName = "RoleMatrix"
Write-Host "Checking $roleListName..." -ForegroundColor Yellow

$roleList = Get-PnPList -Identity $roleListName -ErrorAction SilentlyContinue
if (-not $roleList) {
    Write-Host "Creating list $roleListName..." -ForegroundColor Green
    $roleList = New-PnPList -Title $roleListName -Template GenericList -EnableVersioning
}

# Add Admin column (Multi-person)
$adminField = Get-PnPField -List $roleListName -Identity "Admin" -ErrorAction SilentlyContinue
if (-not $adminField) {
    Write-Host "Adding Admin person column to $roleListName..." -ForegroundColor Green
    Add-PnPField -List $roleListName -Type User -DisplayName "Admin" -InternalName "Admin" -AddToDefaultView
    # Enable multiple selections
    Set-PnPField -List $roleListName -Identity "Admin" -Values @{ AllowMultipleValues = $true }
}

# Seed Application Roles item if missing
$appRoleItem = Get-PnPListItem -List $roleListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Application Roles</Value></Eq></Where></Query></View>"
if (-not $appRoleItem) {
    Write-Host "Seeding 'Application Roles' item..." -ForegroundColor Cyan
    Add-PnPListItem -List $roleListName -Values @{ "Title" = "Application Roles" } | Out-Null
}

# -------------------------------------------------------------------
# 3. MealOptOutRequest List
# -------------------------------------------------------------------
$optOutListName = "MealOptOutRequest"
Write-Host "Checking $optOutListName..." -ForegroundColor Yellow

$optOutList = Get-PnPList -Identity $optOutListName -ErrorAction SilentlyContinue
if (-not $optOutList) {
    Write-Host "Creating list $optOutListName..." -ForegroundColor Green
    $optOutList = New-PnPList -Title $optOutListName -Template GenericList -EnableVersioning
    Set-PnPList -Identity $optOutListName -FieldTitle "Request Summary"
}

# Add Location Lookup Column
$locLookupField = Get-PnPField -List $optOutListName -Identity "Location" -ErrorAction SilentlyContinue
if (-not $locLookupField) {
    Write-Host "Adding Location Lookup to $optOutListName..." -ForegroundColor Green
    Add-PnPField -List $optOutListName -Type Lookup -DisplayName "Location" -InternalName "Location" -AddToDefaultView -LookupList $locationListName -LookupField "Title"
}

# Add MealType Choice Column
$mealTypeField = Get-PnPField -List $optOutListName -Identity "MealType" -ErrorAction SilentlyContinue
if (-not $mealTypeField) {
    Write-Host "Adding MealType choice to $optOutListName..." -ForegroundColor Green
    Add-PnPField -List $optOutListName -Type Choice -DisplayName "Meal Type" -InternalName "MealType" -AddToDefaultView -Choices "Breakfast", "Lunch", "Dinner"
}

# Add MealDate Date Column (Date Only)
$mealDateField = Get-PnPField -List $optOutListName -Identity "MealDate" -ErrorAction SilentlyContinue
if (-not $mealDateField) {
    Write-Host "Adding MealDate to $optOutListName..." -ForegroundColor Green
    Add-PnPField -List $optOutListName -Type DateTime -DisplayName "Meal Date" -InternalName "MealDate" -AddToDefaultView
    Set-PnPField -List $optOutListName -Identity "MealDate" -Values @{ DisplayFormat = 0 } # 0 = DateOnly
}

# Add Requester Person Column (Single)
$requesterField = Get-PnPField -List $optOutListName -Identity "Requester" -ErrorAction SilentlyContinue
if (-not $requesterField) {
    Write-Host "Adding Requester person column to $optOutListName..." -ForegroundColor Green
    Add-PnPField -List $optOutListName -Type User -DisplayName "Requester" -InternalName "Requester" -AddToDefaultView
}

Write-Host "All 3 SharePoint Lists provisioned successfully!" -ForegroundColor Green
