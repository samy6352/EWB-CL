# Console presentation for the installer. Presentation only: nothing here changes what is installed.
# Rich output is used on an interactive terminal. Plain text is kept when output is redirected,
# NO_COLOR is set or CLOUDLENS_PLAIN=1, so logs and automated runs read exactly as before.
$script:UiFancy = ($env:CLOUDLENS_FANCY -eq '1') -or
    (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR -and $env:CLOUDLENS_PLAIN -ne '1')
if ($script:UiFancy) { try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { } }

function Get-UiWidth {
    try { return [math]::Max(60, [math]::Min(96, [Console]::WindowWidth - 4)) } catch { return 80 }
}

function Get-UiBar([int]$Percent, [int]$Width = 30) {
    $filled = [math]::Max(0, [math]::Min($Width, [int][math]::Round($Width * $Percent / 100)))
    return ('█' * $filled) + ('░' * ($Width - $filled))
}

function Show-UiBox([string[]]$Lines, [string]$Color = 'Cyan', [string[]]$Highlight = @()) {
    $w = Get-UiWidth
    $inner = $w - 2
    Write-Host ('  ╭' + ('─' * $inner) + '╮') -ForegroundColor $Color
    foreach ($line in $Lines) {
        $text = if ($line.Length -gt $inner - 2) { $line.Substring(0, $inner - 3) + '…' } else { $line }
        $tone = if ($Highlight -contains $line) { 'White' } else { $Color }
        Write-Host '  │' -ForegroundColor $Color -NoNewline
        Write-Host (' ' + $text.PadRight($inner - 1)) -ForegroundColor $tone -NoNewline
        Write-Host '│' -ForegroundColor $Color
    }
    Write-Host ('  ╰' + ('─' * $inner) + '╯') -ForegroundColor $Color
}

function Show-UiBanner([string]$Subtitle) {
    Write-Host ''
    Show-UiBox -Lines @(
        '',
        '   ☁  C L O U D L E N S',
        '   See. Analyze. Optimize.',
        '',
        "   $Subtitle",
        ''
    ) -Color Cyan -Highlight @('   ☁  C L O U D L E N S')
}

function Write-UiStep([string]$Title, [string]$Label, [int]$Percent, [string]$Elapsed = '') {
    $w = Get-UiWidth
    $left = "  ▸ $Label  $Title"
    $right = if ($Elapsed) { "⏱ $Elapsed" } else { '' }
    $pad = [math]::Max(1, $w + 2 - $left.Length - $right.Length)
    Write-Host ''
    Write-Host $left -ForegroundColor White -NoNewline
    Write-Host (' ' * $pad) -NoNewline
    Write-Host $right -ForegroundColor DarkGray
    Write-Host ('  ▕' + (Get-UiBar $Percent ([math]::Min(40, $w - 14))) + ('▏ {0,3}%' -f $Percent)) -ForegroundColor Cyan
}

function Write-UiStatus([ValidateSet('ok', 'warn', 'bad', 'info')][string]$Kind, [string]$Text) {
    switch ($Kind) {
        'ok'   { Write-Host '   ✔ ' -ForegroundColor Green -NoNewline;  Write-Host $Text }
        'warn' { Write-Host '   ▲ ' -ForegroundColor Yellow -NoNewline; Write-Host $Text -ForegroundColor Yellow }
        'bad'  { Write-Host '   ✖ ' -ForegroundColor Red -NoNewline;    Write-Host $Text -ForegroundColor Red }
        'info' { Write-Host "     $Text" -ForegroundColor DarkGray }
    }
}

function Write-UiSection([string]$Text) {
    Write-Host ''
    Write-Host "  ◆ $Text" -ForegroundColor Cyan
}

function Show-UiSuccess([string]$Url, [string]$SetupUrl, [string]$Monitoring, [string]$Elapsed) {
    Write-Host ''
    $lines = @('', '   ✔  CloudLens is installed', '')
    if ($Url)        { $lines += "   Open         $Url" }
    if ($SetupUrl)   { $lines += "   Track setup  $SetupUrl" }
    if ($Monitoring) { $lines += "   Monitoring   $Monitoring" }
    if ($Elapsed)    { $lines += "   Installed in $Elapsed" }
    $lines += ''
    Show-UiBox -Lines $lines -Color Green -Highlight @('   ✔  CloudLens is installed')
}
