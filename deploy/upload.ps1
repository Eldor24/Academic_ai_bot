# Talaba_Bot'ni Oracle serverga yuklash.
# Oddiy yangilash (serverdagi users.json va .env TEGILMAYDI):
#   powershell -ExecutionPolicy Bypass -File deploy\upload.ps1 -Key "$HOME\.ssh\oracle.key"
# Birinchi o'rnatish yoki .env ni ham yangilash:      ... -WithEnv
# Lokal users.json bilan serverdagini ALMASHTIRISH:    ... -WithUsers   (ehtiyot bo'ling!)
# Serverdagi users.json ni kompyuterga olish:          ... -GetUsers
param(
  [Parameter(Mandatory=$true)][string]$Key,
  [string]$Server = "opc@139.185.55.79",
  [switch]$WithEnv,
  [switch]$WithUsers,
  [switch]$GetUsers
)
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)   # bot papkasi

if ($GetUsers) {
  $dst = "users_server_{0}.json" -f (Get-Date -Format "yyyy-MM-dd_HHmm")
  scp -i $Key "${Server}:/home/opc/talaba_bot/users.json" $dst
  "Serverdagi users.json olindi: $dst"
  exit 0
}

$tar = Join-Path $env:TEMP "talaba_bot.tar.gz"
if (Test-Path $tar) { Remove-Item $tar -Force }
$ex = @("--exclude=venv", "--exclude=.git", "--exclude=__pycache__", "--exclude=eski_versiya",
        "--exclude=Claude outputs", "--exclude=Baza_Zaxira", "--exclude=deploy",
        "--exclude=*.log", "--exclude=*.bak", "--exclude=*.bat",
        "--exclude=*.key", "--exclude=*.pem", "--exclude=*.ppk", "--exclude=*.tar.gz",
        "--exclude=users_*.json", "--exclude=oak_cache.json",
        "--exclude=serverga yuklash pdfni.txt")
if (-not $WithUsers) { $ex += "--exclude=users.json" }
if (-not $WithEnv)   { $ex += "--exclude=.env" }
tar.exe -czf $tar @ex .
"Arxiv: {0:N1} MB  (users.json: {1}, .env: {2})" -f ((Get-Item $tar).Length / 1MB), $(if ($WithUsers) {"BOR"} else {"yo'q"}), $(if ($WithEnv) {"BOR"} else {"yo'q"})
scp -i $Key $tar "$PSScriptRoot\talaba_server_setup.sh" "${Server}:~/"
ssh -i $Key $Server "bash ~/talaba_server_setup.sh"
Remove-Item $tar -Force
