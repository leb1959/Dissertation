# reset_and_upload.ps1
# Ochyschaie repozytoriy i zavantazhuie vsi faily z nulia

$RepoPath  = "E:\КЛОД\Розіл4"
$RemoteURL = "https://github.com/leb1959/Dissertation.git"

$Files = @(
    "Дисертація_Розділи_1_5_v2.docx",
    "Розділ_4_v18.docx",
    "build_Таблиця_4_6.py",
    "build_Таблиця_4_7.py",
    "build_Таблиця_4_8.py",
    "build_Таблиця_4_9.py",
    "build_Таблиця_4_10.py",
    "build_Таблиця_4_12.py",
    "build_Таблиця_4_13.py",
    "build_Додаток_4_3.py",
    "verify_dissertation_numbers.py",
    "run_tukey.py",
    "run_bootstrap.py",
    "run_ks_test.py",
    "Таблиця_4_6.xlsx",
    "Таблиця_4_7.xlsx",
    "Таблиця_4_8.xlsx",
    "Таблиця_4_9.xlsx",
    "Таблиця_4_10.xlsx",
    "Таблиця_4_12.xlsx",
    "Таблиця_4_13.xlsx",
    "tukey_results.xlsx",
    "bootstrap_results.xlsx",
    "ks_test_results.xlsx",
    "Додаток_4_1.xlsx",
    "Додаток_4_2.xlsx",
    "Додаток_4_3.xlsx",
    "Додаток_4_4.xlsx",
    "Додаток_4_5.xlsx",
    "Додаток_4_6.xlsx",
    "Додаток_4_7.xlsx",
    "Додаток_4_8.xlsx",
    "requirements.txt",
    "README.md",
    "СТАН_РОБОТИ.md"
)

# Add git to PATH (GitHub Desktop)
$env:PATH += ";C:\Users\i\AppData\Local\GitHubDesktop\app-3.5.8\resources\app\git\cmd"

Write-Host "=== POVNE PEREZAVANTAZHENNIA ===" -ForegroundColor Cyan
Set-Location $RepoPath

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "POMYLKA: git ne vstanovleno." -ForegroundColor Red
    pause; exit 1
}

if (-not (Test-Path ".git")) {
    git init
    git remote add origin $RemoteURL
    git fetch origin
    git checkout -b main origin/main 2>$null
    if ($LASTEXITCODE -ne 0) { git checkout -b main }
} else {
    $remote = git remote get-url origin 2>$null
    if ($remote -ne $RemoteURL) { git remote set-url origin $RemoteURL }
    git fetch origin
    git checkout main 2>$null
    git reset --hard origin/main 2>$null
}

Write-Host "Vydaliaemo vsi faily..." -ForegroundColor Yellow
git rm -r --cached . -f 2>&1 | Out-Null

Write-Host "Dodaemo faily:" -ForegroundColor Cyan
$added = 0
foreach ($file in $Files) {
    if (Test-Path $file) {
        git add $file
        Write-Host "  + $file" -ForegroundColor Green
        $added++
    } else {
        Write-Host "  - $file (ne znaydeno)" -ForegroundColor Yellow
    }
}
Write-Host "Dodano: $added fayliv" -ForegroundColor Cyan

$date = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "RESET: povne perezavantazhennia ($date)"

Write-Host "Push na GitHub..." -ForegroundColor Cyan
git push origin main --force

if ($LASTEXITCODE -eq 0) {
    Write-Host "=== USPISHNO ===" -ForegroundColor Green
    Write-Host "https://github.com/leb1959/Dissertation" -ForegroundColor Cyan
} else {
    Write-Host "POMYLKA pry push." -ForegroundColor Red
}
pause
