# ============================================================
#  upload_to_github.ps1
#  Завантаження нових/оновлених файлів до репозиторію
#  https://github.com/leb1959/Dissertation
#
#  ЗАПУСК: правою кнопкою → "Run with PowerShell"
#  АБО у PowerShell: cd E:\КЛОД\Розіл4 && .\upload_to_github.ps1
#
#  ВИМОГА: git має бути встановлений та налаштований
# ============================================================

$RepoPath  = "E:\КЛОД\Розіл4"   # шлях до локального репо (якщо клоновано)
$RemoteURL = "https://github.com/leb1959/Dissertation.git"

# --- Файли для завантаження ---
$Files = @(
    # === ФІНАЛЬНІ ДОКУМЕНТИ ===
    "Дисертація_Розділи_1_5_v2.docx",  # ПОВНА дисертація розділи 1-5 (17.05.2026)
    "Розділ_4_v18.docx",               # Розділ 4 окремо, виправлені посилання (17.05.2026)

    # === BUILD-СКРИПТИ ===
    "build_Таблиця_4_6.py",
    "build_Таблиця_4_7.py",
    "build_Таблиця_4_8.py",
    "build_Таблиця_4_9.py",
    "build_Таблиця_4_10.py",
    "build_Таблиця_4_12.py",       # оновлено: T,WS-TZA_Post3; RH,P-pan2
    "build_Таблиця_4_13.py",       # НОВИЙ: Тюкі по місяцях
    "build_Додаток_4_3.py",
    "verify_dissertation_numbers.py",
    "run_tukey.py",
    "run_bootstrap.py",
    "run_ks_test.py",

    # === XLSX ТАБЛИЦІ ===
    "Таблиця_4_6.xlsx",
    "Таблиця_4_7.xlsx",
    "Таблиця_4_8.xlsx",
    "Таблиця_4_9.xlsx",
    "Таблиця_4_10.xlsx",
    "Таблиця_4_12.xlsx",           # оновлено
    "Таблиця_4_13.xlsx",           # НОВИЙ
    "tukey_results.xlsx",
    "bootstrap_results.xlsx",
    "ks_test_results.xlsx",

    # === ДОДАТКИ ===
    "Додаток_4_1.xlsx",
    "Додаток_4_2.xlsx",
    "Додаток_4_3.xlsx",
    "Додаток_4_4.xlsx",
    "Додаток_4_5.xlsx",
    "Додаток_4_6.xlsx",
    "Додаток_4_7.xlsx",
    "Додаток_4_8.xlsx",

    # === СЛУЖБОВІ ===
    "requirements.txt",
    "README.md",
    "СТАН_РОБОТИ.md"
)

Write-Host "=== Завантаження до GitHub ===" -ForegroundColor Cyan

Set-Location $RepoPath

# Перевіряємо чи є git
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "ПОМИЛКА: git не встановлено. Завантажте з https://git-scm.com" -ForegroundColor Red
    pause; exit 1
}

# Ініціалізуємо репо якщо потрібно
if (-not (Test-Path ".git")) {
    Write-Host "Ініціалізація git репозиторію..." -ForegroundColor Yellow
    git init
    git remote add origin $RemoteURL
    git fetch origin
    git checkout -b main origin/main 2>$null
    if ($LASTEXITCODE -ne 0) { git checkout -b main }
}

# Перевіряємо remote
$remote = git remote get-url origin 2>$null
if ($remote -ne $RemoteURL) {
    Write-Host "Встановлюємо remote origin..." -ForegroundColor Yellow
    git remote set-url origin $RemoteURL
}

# Отримуємо останні зміни
Write-Host "Отримуємо останні зміни з GitHub..." -ForegroundColor Yellow
git pull origin main --rebase 2>&1

# Додаємо файли
Write-Host "`nДодаємо файли:" -ForegroundColor Cyan
foreach ($file in $Files) {
    if (Test-Path $file) {
        git add $file
        Write-Host "  + $file" -ForegroundColor Green
    } else {
        Write-Host "  - $file (не знайдено, пропускаємо)" -ForegroundColor Yellow
    }
}

# Перевіряємо чи є що комітити
$status = git status --porcelain
if (-not $status) {
    Write-Host "`nНемає змін для завантаження." -ForegroundColor Yellow
    pause; exit 0
}

# Commit
$date    = Get-Date -Format "yyyy-MM-dd HH:mm"
$message = "Дисертація v2: виправлено посилання GitHub, повний документ розділи 1-5 ($date)"
git commit -m $message
Write-Host "`nКоміт: $message" -ForegroundColor Green

# Push
Write-Host "Завантажуємо на GitHub..." -ForegroundColor Cyan
git push origin main

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== УСПІШНО завантажено! ===" -ForegroundColor Green
    Write-Host "Перегляд: $RemoteURL" -ForegroundColor Cyan
} else {
    Write-Host "`nПОМИЛКА при push. Можливо потрібна авторизація." -ForegroundColor Red
    Write-Host "Спробуйте: git push origin main" -ForegroundColor Yellow
}

pause
