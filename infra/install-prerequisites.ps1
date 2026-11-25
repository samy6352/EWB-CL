function Install-CloudLensWindowsTool {
    param([string]$Name, [string]$PackageId, [string]$ManualUrl, [switch]$NoToolInstall)

    $hint = "Install $Name through your approved software channel ($ManualUrl), then reopen PowerShell 7 and retry."
    if ($NoToolInstall -or -not $IsWindows -or -not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "$Name is required. $hint On managed machines, Azure Cloud Shell (PowerShell) is an alternative if approved."
    }
    Write-Host "Installing $Name with Windows Package Manager..."
    $PSNativeCommandUseErrorActionPreference = $false
    & winget install --id $PackageId --exact --source winget --accept-package-agreements --accept-source-agreements
    # -1978335189: no newer version to apply; -1978335135: already installed. Both mean the tool is present.
    if ($LASTEXITCODE -notin 0, -1978335189, -1978335135) { throw "Windows Package Manager could not install $Name (exit $LASTEXITCODE). $hint" }
    Update-CloudLensPath
}

function Update-CloudLensPath {
    $paths = @($env:PATH, [Environment]::GetEnvironmentVariable('Path', 'User'),
        [Environment]::GetEnvironmentVariable('Path', 'Machine'))
    $env:PATH = (($paths -split ';' | Where-Object { $_ } | Select-Object -Unique) -join ';')
}

function Initialize-CloudLensAzureCli {
    param([switch]$NoToolInstall)

    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        Install-CloudLensWindowsTool 'Azure CLI' 'Microsoft.AzureCLI' 'https://aka.ms/installazurecliwindows' -NoToolInstall:$NoToolInstall
    }
    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        throw 'Azure CLI is not on PATH. Reopen PowerShell 7 and retry.'
    }
    $PSNativeCommandUseErrorActionPreference = $false
    & az version --only-show-errors | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Azure CLI failed to start. Repair or update it through your approved software channel and retry.' }
}

function Connect-CloudLensAzure {
    param([switch]$NonInteractive)

    $PSNativeCommandUseErrorActionPreference = $false
    & az account get-access-token --only-show-errors --output none 2>$null
    if ($LASTEXITCODE -eq 0) { return }
    $hint = "Run 'az login --tenant <customer-tenant-id>' in the approved customer tenant, then retry. If permitted, use 'az login --use-device-code --tenant <customer-tenant-id>' from a terminal without a browser."
    if ($NonInteractive -or -not [Environment]::UserInteractive) { throw "No usable Azure CLI session. $hint" }
    Write-Host 'Azure sign-in required. Complete sign-in with your approved customer account, then return to this terminal.'
    & az login --only-show-errors --output none
    if ($LASTEXITCODE -ne 0) { throw "Azure sign-in failed. $hint" }
    & az account get-access-token --only-show-errors --output none 2>$null
    if ($LASTEXITCODE -ne 0) { throw "Azure sign-in did not produce a usable session. $hint" }
}

function Find-CloudLensPython {
    $PSNativeCommandUseErrorActionPreference = $false
    # A Python installed after this shell opened is on the machine PATH but not in this session.
    Update-CloudLensPath
    foreach ($name in @('python', 'python3')) {
        $command = Get-Command $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $command) { continue }
        # Don't invoke the Microsoft Store placeholder for Python.
        if ($command.Source -match '[\\/]Microsoft[\\/]WindowsApps[\\/]python3?\.exe$') { continue }
        & $command.Source -c 'import sys; sys.exit(sys.version_info < (3, 11))' 2>$null
        if ($LASTEXITCODE -eq 0) { return $command.Source }
    }
    if (-not $IsWindows) { return }
    # Not on PATH: try the py launcher, the standard install folders, then a Store-installed python.
    $candidates = @()
    $launcher = Get-Command py -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($launcher) {
        $resolved = & $launcher.Source -3 -c 'import sys; print(sys.executable)' 2>$null
        if ($LASTEXITCODE -eq 0 -and $resolved) { $candidates += [string]($resolved | Select-Object -First 1) }
    }
    foreach ($base in @("$env:LOCALAPPDATA\Programs\Python", $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($base -and (Test-Path -LiteralPath $base)) {
            $candidates += Get-ChildItem -LiteralPath $base -Directory -Filter 'Python3*' -ErrorAction SilentlyContinue |
                Sort-Object Name -Descending | ForEach-Object { Join-Path $_.FullName 'python.exe' }
        }
    }
    $candidates += Get-Command python, python3 -CommandType Application -ErrorAction SilentlyContinue |
        Where-Object { $_.Source -match '[\\/]Microsoft[\\/]WindowsApps[\\/]python3?\.exe$' } | ForEach-Object Source
    foreach ($path in $candidates) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $out = & $path -c 'import sys; print("ok" if sys.version_info >= (3, 11) else "old")' 2>$null
        if ($LASTEXITCODE -eq 0 -and ($out | Select-Object -First 1) -eq 'ok') { return $path }
    }
}

function Test-CloudLensNode {
    $PSNativeCommandUseErrorActionPreference = $false
    $npm = if ($IsWindows) { 'npm.cmd' } else { 'npm' }
    if (-not (Get-Command node -ErrorAction SilentlyContinue) -or -not (Get-Command $npm -ErrorAction SilentlyContinue)) { return $false }
    & node -e 'const [major, minor] = process.versions.node.split(".").map(Number); process.exit(major === 20 && minor >= 19 || major === 22 && minor >= 12 || major >= 24 ? 0 : 1)' 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }
    & $npm --version 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Initialize-CloudLensBuildTools {
    param([string]$SourceRoot, [switch]$NoToolInstall)

    $PSNativeCommandUseErrorActionPreference = $false
    & az bicep version 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        if ($NoToolInstall) { throw "Bicep is required. Run 'az bicep install' through your approved network and retry." }
        Write-Host 'Installing Bicep for Azure CLI...'
        & az bicep install | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Bicep installation failed. Check proxy/download policy, or have your administrator run 'az bicep install' in the deployment environment."
        }
        & az bicep version 2>$null | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Bicep was installed but Azure CLI cannot run it. Repair the Bicep installation and retry.' }
    }

    $python = Find-CloudLensPython
    if (-not $python) {
        Install-CloudLensWindowsTool 'Python 3.11+' 'Python.Python.3.13' 'https://www.python.org/downloads/' -NoToolInstall:$NoToolInstall | Out-Host
        $python = Find-CloudLensPython
        if (-not $python) { throw 'Python 3.11+ is not on PATH. Reopen PowerShell 7 and retry.' }
    }
    $prebuilt = Test-Path -LiteralPath (Join-Path $SourceRoot 'web/v2/index.html') -PathType Leaf
    if (-not $prebuilt -and (Test-Path -LiteralPath (Join-Path $SourceRoot 'ui/package.json') -PathType Leaf)) {
        if (-not (Test-CloudLensNode)) {
            Install-CloudLensWindowsTool 'Node.js LTS with npm' 'OpenJS.NodeJS.LTS' 'https://nodejs.org/en/download' -NoToolInstall:$NoToolInstall | Out-Host
            if (-not (Test-CloudLensNode)) { throw 'A compatible Node.js LTS and npm were not found on PATH. Reopen PowerShell 7 and retry.' }
        }
    }
    return @{ Python = $python; SkipUiBuild = $prebuilt }
}
