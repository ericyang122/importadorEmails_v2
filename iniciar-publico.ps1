$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$RuntimeDir = Join-Path $Root "runtime"
$EnvFile = Join-Path $Root ".env"
$LocalNgrok = Join-Path $Root ".tools\ngrok.exe"

New-Item -ItemType Directory -Force -Path $RuntimeDir | Out-Null

if (-not (Test-Path -LiteralPath $EnvFile)) {
    throw "Arquivo .env ausente. Configure APP_PASSWORD e APP_SECRET_KEY antes de publicar."
}

$ngrokUrlLine = Get-Content -LiteralPath $EnvFile |
    Where-Object { $_ -match "^\s*NGROK_URL\s*=" } |
    Select-Object -Last 1
$NgrokUrl = if ($ngrokUrlLine) {
    ($ngrokUrlLine -split "=", 2)[1].Trim()
} else {
    ""
}

$pythonCandidates = @(
    @(
        $env:SIGAVI_PYTHON,
        "C:\Users\eric.yang\AppData\Local\Python\pythoncore-3.14-64\python.exe"
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
)

if (-not $pythonCandidates) {
    throw "Python nao encontrado. Defina SIGAVI_PYTHON com o caminho do python.exe."
}
$Python = $pythonCandidates[0]

$ngrokCommand = Get-Command ngrok -ErrorAction SilentlyContinue
if ($ngrokCommand) {
    $Ngrok = $ngrokCommand.Source
} elseif (Test-Path -LiteralPath $LocalNgrok) {
    $Ngrok = $LocalNgrok
} else {
    throw "ngrok nao encontrado. Instale-o em .tools\ngrok.exe ou no PATH."
}

try {
    $health = Invoke-RestMethod -Uri "http://127.0.0.1:5000/health" -TimeoutSec 2
} catch {
    $health = $null
}

if (-not $health.ok) {
    $appProcess = Start-Process `
        -FilePath $Python `
        -ArgumentList @("-m", "waitress", "--listen=127.0.0.1:5000", "app:app") `
        -WorkingDirectory $Root `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $RuntimeDir "app.stdout.log") `
        -RedirectStandardError (Join-Path $RuntimeDir "app.stderr.log") `
        -PassThru
    Set-Content -LiteralPath (Join-Path $RuntimeDir "app.pid") -Value $appProcess.Id

    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        Start-Sleep -Milliseconds 500
        try {
            $health = Invoke-RestMethod -Uri "http://127.0.0.1:5000/health" -TimeoutSec 2
            if ($health.ok) {
                $ready = $true
                break
            }
        } catch {
        }
    }
    if (-not $ready) {
        throw "O Flask nao iniciou. Consulte runtime\app.stderr.log."
    }
}

try {
    $tunnels = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
} catch {
    $tunnels = $null
}

if (-not $tunnels.tunnels) {
    $ngrokArguments = @("http", "5000", "--log=stdout", "--log-format=json")
    if ($NgrokUrl) {
        $ngrokArguments += @("--url", $NgrokUrl)
    }
    $ngrokProcess = Start-Process `
        -FilePath $Ngrok `
        -ArgumentList $ngrokArguments `
        -WorkingDirectory $Root `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $RuntimeDir "ngrok.stdout.log") `
        -RedirectStandardError (Join-Path $RuntimeDir "ngrok.stderr.log") `
        -PassThru
    Set-Content -LiteralPath (Join-Path $RuntimeDir "ngrok.pid") -Value $ngrokProcess.Id

    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        Start-Sleep -Milliseconds 500
        if ($ngrokProcess.HasExited) {
            throw "O ngrok encerrou. Configure o authtoken e consulte runtime\ngrok.stderr.log."
        }
        try {
            $tunnels = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 2
            if ($tunnels.tunnels) {
                break
            }
        } catch {
        }
    }
}

$publicUrl = $tunnels.tunnels |
    Where-Object { $_.proto -eq "https" } |
    Select-Object -First 1 -ExpandProperty public_url

if (-not $publicUrl) {
    throw "Tunel HTTPS nao ficou disponivel. Consulte os logs em runtime."
}

Set-Content -LiteralPath (Join-Path $RuntimeDir "url.txt") -Value $publicUrl
Write-Host ""
Write-Host "Link publico: $publicUrl"
Write-Host "Senha da ferramenta: consulte APP_PASSWORD no arquivo .env"
