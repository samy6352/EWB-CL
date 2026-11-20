#requires -Version 7.0
[CmdletBinding(DefaultParameterSetName = 'Automatic')]
param(
    [string]$SourceRoot,
    [guid]$SubscriptionId,
    [Parameter(Mandatory)][string]$ResourceGroup,
    [string]$DeploymentName,
    [ValidateNotNullOrEmpty()][string]$Scope,
    [Parameter(Mandatory, ParameterSetName = 'Existing')][guid]$EntraClientId,
    [Parameter(Mandatory, ParameterSetName = 'Create')][switch]$CreateEntraApp,
    [switch]$Apply,
    [Parameter(ParameterSetName = 'Automatic')][switch]$AutoApply
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    throw 'Use an approved PowerShell 7 FullLanguage session.'
}
if ($CreateEntraApp -and -not $Apply) {
    throw 'Creating an Entra app is a write. Use -Apply for the guided creation and recovery, or supply an existing -EntraClientId for preview only.'
}

function Get-AzureJson([string[]]$Arguments) {
    $result = & az @Arguments --only-show-errors -o json
    if ($LASTEXITCODE -ne 0 -or -not $result) { throw "Azure CLI could not complete: $($Arguments[0..1] -join ' ')." }
    return ($result -join "`n") | ConvertFrom-Json
}

$account = Get-AzureJson @('account', 'show')
if (-not $SubscriptionId) { $SubscriptionId = [guid]$account.id }
if ($account.id -ne "$SubscriptionId") { throw "Select the original subscription first: az account set --subscription $SubscriptionId" }
if (-not $Scope) {
    $Scope = "$SubscriptionId"
    Write-Host 'No monitoring scope supplied; recovery will monitor the hosting subscription only.'
}
if ($PSCmdlet.ParameterSetName -eq 'Automatic' -and -not $AutoApply) {
    throw 'Supply -AutoApply to authorise automatic registration and recovery, or -EntraClientId for preview.'
}

function Get-LatestCloudLensDeployment {
    $runs = @(Get-AzureJson @('deployment', 'sub', 'list', '--subscription', "$SubscriptionId"))
    foreach ($run in ($runs | Where-Object name -Match '^cloudlens-\d{8}-\d{6}(?:-r\d+)?$' |
            Sort-Object { [datetimeoffset]$_.properties.timestamp } -Descending)) {
        $targetGroup = $run.properties.parameters.resourceGroupName.value
        if ($targetGroup -and $targetGroup -ne $ResourceGroup) { continue }
        $details = Get-AzureJson @('deployment', 'sub', 'show', '--subscription', "$SubscriptionId", '--name', $run.name)
        if ($details.properties.parameters.resourceGroupName.value -eq $ResourceGroup) { return $details }
    }
    throw "No CloudLens deployment found for '$ResourceGroup' in '$SubscriptionId'. Verify the Azure CLI subscription or specify -DeploymentName."
}

function Assert-NoActiveDeployment {
    $latest = Get-LatestCloudLensDeployment
    if ($latest.name -ne $deployment.name -or $latest.properties.provisioningState -ne 'Failed') {
        throw "The latest CloudLens deployment ($($latest.name)) is $($latest.properties.provisioningState), or no longer matches the selected failed run. Do not start a parallel repair."
    }
    $runs = @(Get-AzureJson @('deployment', 'group', 'list', '--subscription', "$SubscriptionId", '--resource-group', $ResourceGroup))
    if ($runs | Where-Object { $_.properties.provisioningState -notin 'Succeeded', 'Failed', 'Canceled' }) {
        throw 'A deployment is still active in the resource group. Wait for it to finish before repairing.'
    }
}

$deployment = if ($DeploymentName) {
    Get-AzureJson @('deployment', 'sub', 'show', '--subscription', "$SubscriptionId", '--name', $DeploymentName)
} else {
    Get-LatestCloudLensDeployment
}
Assert-NoActiveDeployment
$p = $deployment.properties.parameters
if ($p.resourceGroupName.value -ne $ResourceGroup -or
    $p.deployEntraApp.value -ne $true -or $p.authDelegatedArm.value -ne $false -or
    $p.authAllowLocal.value -ne $false -or $p.managedExports.value -ne $true -or $p.networkIsolation.value -ne $true) {
    throw 'Use the failed subscription deployment for this resource group. This script supports the guided managed-identity installation that attempted to create an Entra app.'
}
if ($AutoApply) {
    $operations = @(Get-AzureJson @('deployment', 'operation', 'group', 'list',
        '--subscription', "$SubscriptionId", '--resource-group', $ResourceGroup, '--name', 'cloudlens-entra'))
    if (-not ($operations | Where-Object {
        $_.properties.provisioningState -eq 'Failed' -and
        ($_.properties.targetResource.resourceType -match '^Microsoft\.Graph/applications(?:@|$)') -and
        (($_.properties.statusMessage | ConvertTo-Json -Depth 20 -Compress) -match 'Authorization_RequestDenied')
    })) {
        throw 'Automatic repair requires a recorded Entra app-registration permission failure in cloudlens-entra.'
    }
}
$site = @(Get-AzureJson @('webapp', 'list', '--subscription', "$SubscriptionId", '--resource-group', $ResourceGroup) |
    Where-Object { $_.tags.application -eq 'CloudLens' })
if ($site.Count -ne 1 -or -not $site[0].defaultHostName) { throw 'Expected exactly one existing CloudLens web app in the resource group.' }
$url = "https://$($site[0].defaultHostName)"
$callback = "$url/auth/callback"

$size = switch ("$($p.appServicePlanSku.value)/$($p.postgresSku.value)/$($p.postgresStorageGb.value)") {
    'B1/Standard_D2ds_v5/64' { 'small' }
    'B2/Standard_D2ds_v5/128' { 'medium' }
    'B2/Standard_E4ds_v5/256' { 'large' }
    default { throw 'Original sizing does not match a guided installer profile. Recover with the original deployment parameters.' }
}
$settings = @{
    SubscriptionId = "$SubscriptionId"
    ResourceGroup = $ResourceGroup
    Scope = $Scope.Trim()
    Size = $size
    AppIdentityOnly = $true
    NonInteractive = $true
    NoToolInstall = $true
}
if ($settings.Scope -notmatch '^(all|(?:/subscriptions/)?[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}|/providers/Microsoft\.Management/managementGroups/[A-Za-z0-9._()-]{1,90})$') {
    throw 'Supply the original monitoring scope (subscription ID, all, or management-group resource ID).'
}
foreach ($key in @('Location', 'AiLocation', 'EntraAppUniqueName', 'NamePrefix', 'NameSuffix',
        'InitialHistoryMonths', 'VnetAddressPrefix', 'AppSubnetAddressPrefix',
        'PostgresSubnetAddressPrefix', 'PrivateEndpointSubnetAddressPrefix')) {
    $property = $p.PSObject.Properties[$key.Substring(0, 1).ToLowerInvariant() + $key.Substring(1)]
    if (-not $property -or $null -eq $property.Value.value) { throw "The failed deployment did not record $key." }
    $settings[$key] = $property.Value.value
}
if ($settings.NamePrefix -eq 'cloudlens' -and -not $settings.NameSuffix) {
    # ARM records the generated-name prefix, but install.ps1 expects both naming overrides to be omitted.
    $settings.Remove('NamePrefix')
    $settings.Remove('NameSuffix')
}
$scopeKey = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
    [Text.Encoding]::UTF8.GetBytes("$SubscriptionId/$ResourceGroup".ToLowerInvariant()))).ToLowerInvariant()
$inventory = @(Get-AzureJson @('resource', 'list', '--subscription', "$SubscriptionId", '--resource-group', $ResourceGroup))
if (-not $p.storageAccountName.value -or -not ($inventory | Where-Object {
        $_.type -eq 'Microsoft.Storage/storageAccounts' -and $_.name -eq $p.storageAccountName.value
    })) {
    throw 'The original storage account was not found in this resource group. Do not start a fresh deployment.'
}
if (-not $SourceRoot) {
    $cloudLensHome = Join-Path $HOME 'CloudLens'
    $sources = @(if (Test-Path -LiteralPath $cloudLensHome -PathType Container) {
        Get-ChildItem -LiteralPath $cloudLensHome -Directory | ForEach-Object {
            $candidate = Join-Path $_.FullName 'source'
            $candidateState = Join-Path $candidate "data\deployments\$scopeKey.storage.json"
            if ((Test-Path -LiteralPath $candidateState -PathType Leaf) -and
                (Test-Path -LiteralPath (Join-Path $candidate 'infra\install.ps1') -PathType Leaf)) {
                $state = Get-Content -LiteralPath $candidateState -Raw | ConvertFrom-Json
                if ($state.storageAccountName -eq $p.storageAccountName.value) { $candidate }
            }
        }
    })
    if ($sources.Count -ne 1) { throw 'Could not identify one retained source folder. Pass -SourceRoot using the path printed by the failed run.' }
    $SourceRoot = $sources[0]
}
$installer = Join-Path $SourceRoot 'infra\install.ps1'
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) { throw 'Use the retained source folder from the failed installation.' }
$stateFile = Join-Path $SourceRoot "data\deployments\$scopeKey.storage.json"
if (-not (Test-Path -LiteralPath $stateFile -PathType Leaf)) { throw 'The original storage naming state is missing. Use the retained source folder from the failed run.' }
$storageState = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
if (-not $p.storageAccountName.value -or $storageState.storageAccountName -ne $p.storageAccountName.value) {
    throw 'The retained storage name does not match the failed deployment.'
}

Write-Host "Found $($inventory.Count) existing resource(s) in $ResourceGroup. Recovering $($deployment.name)."
Write-Host "Existing app: $url | Monitoring scope: $Scope"
if ($AutoApply -or $CreateEntraApp) {
    $registrationReceipt = "$stateFile.entra-repair.json"
    if ($AutoApply -and (Test-Path -LiteralPath $registrationReceipt -PathType Leaf)) {
        $receipt = Get-Content -LiteralPath $registrationReceipt -Raw | ConvertFrom-Json
        if ($receipt.tenantId -ne $account.tenantId -or $receipt.callback -cne $callback) {
            throw 'The previously created registration receipt does not match this tenant and app.'
        }
        $EntraClientId = [guid]$receipt.clientId
    } else {
        $helper = Join-Path $PSScriptRoot 'New-CloudLensEntraApp.ps1'
        if (-not (Test-Path -LiteralPath $helper -PathType Leaf)) { throw 'Keep New-CloudLensEntraApp.ps1 next to this repair script.' }
        if (-not $AutoApply -and (Read-Host "Type CREATE ENTRA to register the sign-in app in tenant $($account.tenantId)") -cne 'CREATE ENTRA') {
            Write-Host 'Cancelled before creating a registration.'
            return
        }
        Assert-NoActiveDeployment
        $registration = & $helper -TenantId $account.tenantId -DisplayName "CloudLens - $ResourceGroup" -CallbackUri $callback
        $EntraClientId = [guid]$registration.EntraClientId
        @{ tenantId = $account.tenantId; callback = $callback; clientId = "$EntraClientId" } |
            ConvertTo-Json | Set-Content -LiteralPath $registrationReceipt
        Write-Host "Entra application client ID: $EntraClientId (reuse this ID if you need to resume again)"
    }
}
$registration = Get-AzureJson @('ad', 'app', 'show', '--id', "$EntraClientId")
if ($registration.signInAudience -ne 'AzureADMyOrg' -or $callback -cnotin @($registration.publicClient.redirectUris)) {
    throw "Ask the Entra admin to configure $callback as a Mobile and desktop applications redirect URI on the approved single-tenant registration."
}
if (-not (Get-AzureJson @('ad', 'sp', 'show', '--id', "$EntraClientId")).id) {
    throw 'The Entra enterprise application could not be verified.'
}
$settings.EntraClientId = "$EntraClientId"

& $installer @settings -WhatIf
if (-not $?) { throw 'Recovery What-If failed. Do not apply.' }
if (-not $Apply -and -not $AutoApply) { return }
if (-not $AutoApply -and (Read-Host "Review the changes, then type RESUME $ResourceGroup to reconcile the existing deployment and finish setup") -cne "RESUME $ResourceGroup") {
    Write-Host 'Deployment not resumed. Any Entra registration created in this session was retained.'
    return
}
Assert-NoActiveDeployment
& $installer @settings
if (-not $?) { throw 'Recovery did not complete. Inspect the error before rerunning.' }
$ready = Invoke-RestMethod -Uri "$url/readyz" -TimeoutSec 30
if ($ready.status -ne 'ready') { throw "The existing app did not report ready: $url/readyz" }
Write-Host "CloudLens is ready: $url. Verify Microsoft sign-in and review Setup for export status."
