#requires -Version 5.1
<#
.SYNOPSIS
    Автономний CLI-асистент PS-COPILOT для PowerShell 5.1 (Multi-LLM).
.DESCRIPTION
    Цей скрипт надає інтерактивний інтерфейс термінала для спілкування з
    Gemini (або іншими LLM через API, сумісний з OpenAI). Він підтримує збереження історії,
    завантаження та експорт у форматі markdown, адаптований для обмежених середовищ Windows 11.
.AUTHOR
    Роман Падалко 
.DATE
    Липень 2026
#>

<#
Оскільки це скрипт, Windows за замовчуванням блокує його запуск з міркувань безпеки.
Щоб дозволити роботу скриптів для поточного користувача, скопіюйте наступну команду, вставте її в вікно PowerShell і натисніть Enter:
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Якщо система запитає підтвердження, натисніть "Y" (Так) та Enter.
#>

# ===================================================================
# --- Початкове налаштування середовища ---

# Примусове використання TLS 1.2 або TLS 1.3 для захисту зв'язку з API (критично для PS 5.1)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

# Налаштування проксі (щоб уникнути помилок в корпоративній мережі)
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials

# Налаштування поточної сесії консолі для правильного прийняття та відображення українських символів
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ===================================================================
# --- Конфігурація та шляхи ---

# Динамічне визначення каталогу скрипта
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = $pwd.Path }

# Автоматичне створення папок якщо вони відсутні
$HistoryDir = Join-Path $ScriptDir "chats_history"
if (-not (Test-Path $HistoryDir)) {
    New-Item -ItemType Directory -Path $HistoryDir -Force | Out-Null
}
$ExportDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportDir)) { 
    New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null 
}

# Значення за замовчуванням для OpenAI-сумісного ендпоінту Gemini API
$currentKey = ""
$ApiUrl = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
$ModelName = "gemini-3.6-flash"
# $ModelName = "gemini-3.7-flash"
# $ModelName = "gemini-flash-latest"
# $ModelName = "gemini-pro-latest"

$SystemPrompt = @'
Ти — високоінтелектуальний універсальний CLI-асистент PS-COPILOT. Твоя мета — надавати якісну, точну та структуровану допомогу з будь-яких тем: від аналізу даних, програмування та адміністрування до генерації текстів, брейнштормінгу, навчання та повсякденних консультацій.

ГОЛОВНІ ПРАВИЛА СПІЛКУВАННЯ:
1. Мова та стиль:
   - Відповідай виключно українською мовою.
   - Адаптуй тон під контекст: для технічних/наукових питань надавай глибокі, аналітичні відповіді; для творчих або побутових — зрозумілі, лаконічні та дружні.
   - Уникай "води", поверхневих пояснень та зайвих вступів. Одразу переходь до суті.

2. Принцип хірургічної точності (При роботі з текстом або кодом користувача):
   - Якщо користувач надає свій текст, документ або код для редагування, аналізу чи виправлення — змінюй ТІЛЬКИ ті блоки, про які йдеться у запиті.
   - Не роби несанкціонованого рефакторингу чи спрощення. Зберігай стиль, структуру, авторські коментарі та налагоджену логіку.
   - При зміні коду чітко виділяй нові або змінені рядки коментарями (наприклад: `# UPDATED: ...` або `# CHANGED: ...`).

3. Робота з кодом та автоматизацією (якщо є у запиті):
   - Завжди обгортай код у відповідні блоки із зазначенням мови (```powershell, ```python, ```bash, ```json тощо).
   - Додавай зрозумілі коментарі до складних або критичних ділянок коду.
   - Обов'язково та зрозуміло пояснюй, що робить наведений код та як він працює.
   - Окремо та явно попереджай про потенційні ризики (наприклад: видалення даних, зміна системних налаштувань, навантаження на мережу чи потенційні витрати).
   - Якщо просять код PowerShell, за замовчуванням орієнтуйся на сумісність із Windows PowerShell 5.1 та Windows 11 (уникай специфічного синтаксису PS 7+, якщо про це не попросили явно).

4. Форматування для CLI-Термінала (Dracula ANSI Engine):
   - Активно використовуй заголовки (`#`, `##`, `###`) для візуального розділення блоків думки.
   - Для переліків, алгоритмів або чеклістів використовуй списки та інтерактивні чекбокси: `- [ ]` (невиконано) та `- [x]` (виконано).
   - Для порівняння даних, характеристик чи параметрів використовуй Markdown-таблиці.
   - Виділяй ключові думки **жирним шрифтом**, а терміни, команди чи шляхи — `інлайн-кодом`.
   - Важливі примітки та застереження обов'язково оформлюй у цитати:
     > 💡 **Примітка:** Корисна порада або важливий нюанс.
     > ⚠️ **Увага (Ризик):** Застереження про деструктивні дії, видалення чи системні зміни.
'@

# Ініціалізація змінних сесії
$script:ChatHistory = @()
$script:CurrentSessionFile = ""
$script:IsModified = $false

# ===================================================================
# ЦЕНТРАЛІЗОВАНА ПАЛІТРА DRACULA (ANSI TrueColor 24-bit)
# ===================================================================
$esc = [char]27
$script:UI = @{
    Reset       = "$esc[0m"
    Bold        = "$esc[1m"
    Italic      = "$esc[3m"
    Underline   = "$esc[4m"

    # Dracula Palette
    BgSelection = "$esc[48;2;68;71;90m"         # Selection Background
    Fg          = "$esc[38;2;248;248;242m"      # Main Text (Soft White)
    Comment     = "$esc[38;2;98;114;164m"       # Muted Blue-Gray (Borders/Muted)
    Pink        = "$esc[38;2;255;121;198m"      # Primary Accents / H1
    Purple      = "$esc[38;2;189;147;249m"      # H2 / Secondary Accents
    Cyan        = "$esc[38;2;139;233;253m"      # User / H3 / Links
    Green       = "$esc[38;2;80;250;123m"       # AI / Success / Strings
    Yellow      = "$esc[38;2;241;250;140m"      # Warnings / Numbers / Bullets
    Orange      = "$esc[38;2;255;184;108m"      # Commands / Code Inline
    Red         = "$esc[38;2;255;85;85m"        # Errors
}

# ДОПОМІЖНІ ФУНКЦІЇ ДЛЯ СИСТЕМНИХ СПОВІЩЕНЬ (DRACULA STYLE)

function Write-Success ([string]$Text, [switch]$NoNewline) {
    $msg = "$($script:UI.Green)✅ $Text$($script:UI.Reset)"
    if ($NoNewline) { Write-Host $msg -NoNewline } else { Write-Host $msg }
}

function Write-ErrorMsg ([string]$Text, [switch]$NoNewline) {
    $msg = "$($script:UI.Red)❌ $Text$($script:UI.Reset)"
    if ($NoNewline) { Write-Host $msg -NoNewline } else { Write-Host $msg }
}

function Write-Warn ([string]$Text, [switch]$NoNewline) {
    $msg = "$($script:UI.Yellow)⚠️ $Text$($script:UI.Reset)"
    if ($NoNewline) { Write-Host $msg -NoNewline } else { Write-Host $msg }
}

function Write-Info ([string]$Text, [switch]$NoNewline) {
    $msg = "$($script:UI.Cyan)ℹ️ $Text$($script:UI.Reset)"
    if ($NoNewline) { Write-Host $msg -NoNewline } else { Write-Host $msg }
}

function Write-Muted ([string]$Text, [switch]$NoNewline) {
    $msg = "$($script:UI.Comment)$Text$($script:UI.Reset)"
    if ($NoNewline) { Write-Host $msg -NoNewline } else { Write-Host $msg }
}

function Write-Line {
    Write-Muted "$([string]::new('─', 78))"
}

# ===================================================================
# --- Функції керування сесією ---
function Reset-Session {
    $script:ChatHistory = @(
        @{ role = "system"; content = $SystemPrompt }
    )
    $script:CurrentSessionFile = ""
	$script:IsModified = $false
}

function Get-ApiKey {
    [CmdletBinding()]
    param (
        # Ім'я файлу для збереження зашифрованого ключа
        [string]$FileName = "api_key",

        # Прапорець для примусового запиту та перезапису ключа
        [switch]$Force
    )

    # 1. Визначення каталогу скрипта
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $pwd.Path }
    $keyFilePath = Join-Path -Path $scriptDir -ChildPath $FileName

    # 2. Спроба зчитати існуючий ключ (ТІЛЬКИ якщо НЕ вказано параметр -Force)
    if ((Test-Path -Path $keyFilePath) -and -not $Force) {
        try {
            $encryptedData = Get-Content -Path $keyFilePath -Raw
            if (-not [string]::IsNullOrWhiteSpace($encryptedData)) {
                # Розшифровуємо через DPAPI
                $secureKey = ConvertTo-SecureString -String $encryptedData.Trim()
                $plainKey = (New-Object System.Net.NetworkCredential("", $secureKey)).Password

                if (-not [string]::IsNullOrWhiteSpace($plainKey)) {
                    return $plainKey
                }
            }
        }
        catch {
            Write-Warn "Не вдалося розшифрувати збережений API ключ (файл пошкоджено або змінено користувача)."
            Remove-Item -Path $keyFilePath -Force -ErrorAction SilentlyContinue
        }
    }

    # 3. Інформування користувача про причину запиту ключа
    if ($Force) {
        Write-Warn "🔄 Викликано примусову зміну API-ключа (параметр -Force)."
    } else {
        Write-Warn "🔑 API ключ не знайдено."
    }

    Write-Host ""
    Write-Warn " Введіть новий API Key (символи приховано): " -NoNewline

    # 4. Введення нового ключа та маскування символів
    $secureEnteredKey = Read-Host -AsSecureString
    $plainEnteredKey = (New-Object System.Net.NetworkCredential("", $secureEnteredKey)).Password

    # 5. Валідація, зашифрування та перезапис
    if (-not [string]::IsNullOrWhiteSpace($plainEnteredKey)) {
        try {
            # Перетворюємо SecureString на зашифрований текст DPAPI
            $encryptedToSave = ConvertFrom-SecureString -SecureString $secureEnteredKey

            # Перезаписуємо або створюємо файл із ключем
            [System.IO.File]::WriteAllText($keyFilePath, $encryptedToSave, [System.Text.Encoding]::UTF8)

            Write-Success "Новий API-ключ успішно зашифровано та збережено у $FileName"
            return $plainEnteredKey.Trim()
        }
        catch {
            Write-ErrorMsg "Помилка збереження зашифрованого ключа: $_"
            throw
        }
    }
    else {
        Write-ErrorMsg "Помилка: API Ключ не вказано. Операцію скасовано."
        throw
    }
}

function Show-WelcomeScreen {
    $c = $script:UI
    Clear-Host

    # 1. Фірмовий ASCII-логотип PS-COPILOT (Dracula Yellow + Pink)
    Write-Host ""
    Write-Host "$($c.Pink)  █▀█ █▀ $($c.Yellow)  █▀▀ █▀█ █▀█ █ █  █▀█ ▀█▀$($c.Reset)"
    Write-Host "$($c.Pink)  █▀▀ ▄█ $($c.Yellow)  █▄▄ █▄█ █▀▀ █ █▄ █▄█  █ $($c.Reset)"
    Write-Host "$($c.Comment)   ┘   ─┘   ──┘ ──┘ ┘   ┘ ─┘ ──┘  ┘ $($c.Reset)"
    Write-Host ""
    # 2. Інформаційна панель агента (Dracula Comment / Cyan / Green)
    Write-Host "$($c.Comment)┌────────────────────────────────────────────────────────────────────────────┐$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)  $($c.Pink)$($c.Bold)🛡️ PS-COPILOT CORE$($c.Reset) | $($c.Cyan)PowerShell Native Intelligent Copilot   $($c.Reset)             $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)├────────────────────────────────────────────────────────────────────────────┤$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)  Модель : $($c.Green)$ModelName$($c.Reset) $($c.Comment)(OpenAI-Compatible API)$($c.Reset)                         $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)  Сесія  : $($c.Comment)./chats_history$($c.Reset)                                                  $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)├────────────────────────────────────────────────────────────────────────────┤$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)  $($c.Yellow)Системні команди:$($c.Reset)                                                         $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/new$($c.Reset)     - Очистити історію поточної сесії                               $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/save$($c.Reset)    - Зберегти діалог у файл JSON                                   $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/load$($c.Reset)    - Завантажити діалог із файлу JSON                              $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/export$($c.Reset)  - Експортувати діалог у MD                                      $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/refresh$($c.Reset) - Очистити екран та оновити діалог                              $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/p$($c.Reset)       - Вставити вміст буфера обміну в промпт                         $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/c$($c.Reset)       - Копіювати відповідь до буфера обміну                          $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/key$($c.Reset)     - Оновити/замінити API-ключ                                     $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/help,/?$($c.Reset) - Показати це вікно допомоги                                    $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)│$($c.Reset)   $($c.Orange)/exit$($c.Reset)    - Закрити програму                                              $($c.Comment)│$($c.Reset)"
    Write-Host "$($c.Comment)└────────────────────────────────────────────────────────────────────────────┘$($c.Reset)"
}

function Refresh-ChatHistory {
    $c = $script:UI

    Write-Info "ІСТОРІЮ ДІАЛОГУ ОНОВЛЕНО "
    Write-Line

    if ($script:ChatHistory.Count -le 1) {
        Write-Muted "Історія діалогу порожня."
        return
    }

    foreach ($msg in $script:ChatHistory) {
        if ($msg.role -eq "system") { continue }

        if ($msg.role -eq "user") {
            Write-Host "`n$($c.Green)$($c.Bold)🧑🏻 Користувач:$($c.Reset)"
            Write-Host $msg.content
        }
        elseif ($msg.role -eq "assistant") {
            Write-Host "`n$($c.Cyan)$($c.Bold)🤖 PS-COPILOT [$ModelName]:$($c.Reset)"
            $ansiReply = Convert-MarkdownToAnsi -MarkdownText $msg.content
            Write-Host $ansiReply
        }
    }
    Write-Line
}

function Save-ChatHistory {
    param (
        [string]$FileName = ""
    )
    if ($script:ChatHistory.Count -le 1) {
        Write-Warn "Немає повідомлень для збереження!"
        return
    }

    # Якщо ім'я файлу не передано, але у нас вже є активний файл сесії (завантажений або збережений раніше)
    if ([string]::IsNullOrEmpty($FileName) -and -not [string]::IsNullOrEmpty($script:CurrentSessionFile)) {
        $fileNameOnly = Split-Path $script:CurrentSessionFile -Leaf
        Write-Warn "Зберегти зміни у поточний файл '$fileNameOnly'? (y/n): " -NoNewline
        $confirm = Read-Host
        if ($confirm -eq "y" -or $confirm -eq "yes") {
            $FileName = $fileNameOnly
        }
    }

    if ([string]::IsNullOrEmpty($FileName)) {
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $suggestedName = "chat_$timestamp.json"
        Write-Warn "Введіть ім'я файлу (натисніть Enter для '$suggestedName'): " -NoNewline
        $inputName = Read-Host
        if ([string]::IsNullOrEmpty($inputName)) {
            $FileName = $suggestedName
        } else {
            if (-not $inputName.EndsWith(".json")) { $inputName += ".json" }
            $FileName = $inputName
        }
    }

    $filePath = Join-Path $HistoryDir $FileName
    try {
        # PS 5.1 ConvertTo-Json нативно обробляє хеш-таблиці, якщо вказано параметр -Depth.
        # Ми явно встановлюємо глибину 100, щоб уникнути згортання структури.
        $jsonContent = $script:ChatHistory | ConvertTo-Json -Depth 100

        # Явне використання .NET IO забезпечує коректне кодування UTF-8
        [System.IO.File]::WriteAllText($filePath, $jsonContent, [System.Text.Encoding]::UTF8)

        $script:CurrentSessionFile = $filePath
        $script:IsModified = $false # <--- Скидаємо прапорець змін після успішного збереження

        Write-Success "Історію успішно збережено у: " -NoNewline
        Write-Info $filePath
    }
    catch {
        Write-ErrorMsg "Помилка збереження файлу: $_"
    }
}

function Load-ChatHistory {
    $files = Get-ChildItem -Path $HistoryDir -Filter "*.json"
    if ($files.Count -eq 0) {
        Write-Warn "Збережених діалогів у папці не знайдено: $HistoryDir"
        return
    }

    Write-Info "Доступні збережені діалоги"
    Write-Line
    
    for ($i = 0; $i -lt $files.Count; $i++) {
        $lastWrite = $files[$i].LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        Write-Host "$($script:UI.Yellow)[$i]$($script:UI.Reset) $($files[$i].Name) $($script:UI.Comment)($lastWrite)$($script:UI.Reset)"
    }
    Write-Muted "[c] Скасувати"
    Write-Host ""
    Write-Warn "Виберіть номер файлу для завантаження: " -NoNewline
    $choice = Read-Host
    if ($choice -eq "c" -or [string]::IsNullOrEmpty($choice)) {
        Write-Muted "Завантаження скасовано."
        return
    }

    $index = -1
    if ([int]::TryParse($choice, [ref]$index) -and $index -ge 0 -and $index -lt $files.Count) {
        $targetFile = $files[$index].FullName
        try {
            $jsonRaw = [System.IO.File]::ReadAllText($targetFile, [System.Text.Encoding]::UTF8)
            $parsedHistory = $jsonRaw | ConvertFrom-Json

            # Відновлення історії чату як масиву структурованих хеш-таблиць
            $script:ChatHistory = @()
            foreach ($msg in $parsedHistory) {
                $script:ChatHistory += @{
                    role = $msg.role
                    content = $msg.content
                }
            }

            $script:CurrentSessionFile = $targetFile
			$script:IsModified = $false
			
            Write-Success "Успішно завантажено сесію з: " -NoNewline
            Write-Info $files[$index].Name
            Write-Warn "Відновлено $($script:ChatHistory.Count - 1) повідомлень."
			
			Start-Sleep -Milliseconds 800 # Коротка пауза для читання статусу
        }
        catch {
            Write-ErrorMsg "Помилка завантаження файлу: $_"
        }
    } else {
        Write-ErrorMsg "Невірний вибір."
    }
    Write-Line
}

function Export-ToMarkdown {
    if ($script:ChatHistory.Count -le 1) {
        Write-Warn "Немає діалогу для експорту!"
        return
    }

    # Визначення назви файлу на основі поточного json або timestamp
    if (-not [string]::IsNullOrEmpty($script:CurrentSessionFile)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:CurrentSessionFile)
        $fileName = "$baseName.md"
    } else {
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $fileName = "export_$timestamp.md"
    }

    $filePath = Join-Path $ExportDir $fileName

    # Формування чистого Markdown-документа
    $mdLines = [System.Collections.Generic.List[string]]::new()
    $mdLines.Add("# Експорт діалогу PS-COPILOT ($((Get-Date).ToString('yyyy-MM-dd HH:mm')))`n")

    foreach ($msg in $script:ChatHistory) {
        if ($msg.role -eq "system") { continue }

        if ($msg.role -eq "user") {
            $mdLines.Add("### 🧑🏻 Користувач:`n")
            $mdLines.Add($msg.content)
            $mdLines.Add("`n---`n")
        }
        elseif ($msg.role -eq "assistant") {
            $mdLines.Add("### 🤖 PS-COPILOT:`n")
            $mdLines.Add($msg.content)
            $mdLines.Add("`n---`n")
        }
    }

    try {
        # .NET WriteAllText з [System.Text.Encoding]::UTF8 створює файл у UTF-8-BOM у PS 5.1
        $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($filePath, ($mdLines -join "`n"), $utf8WithBom)

        Write-Success "📝 Диалог успішно експортовано в MD: " -NoNewline
        Write-Info $filePath
    }
    catch {
        Write-ErrorMsg "Помилка експорту в MD: $_"
    }
}

# ===================================================================
# Конвертування виводу MD в термінал
function Convert-MarkdownToAnsi {
    <#
    .SYNOPSIS
        Конвертує MD/GFM текст у консольний формат ANSI TrueColor (Dracula Theme).
        Повністю сумісно з Windows PowerShell 5.1 та PowerShell 7+.
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, Position = 0)]
        [String]$MarkdownText
    )

    begin {
        # 1. Захист середовища: зберігаємо кодування консолі
        $script:originalEncoding = [Console]::OutputEncoding
        if ([Console]::OutputEncoding.EncodingName -ne [System.Text.Encoding]::UTF8.EncodingName) {
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        }

        $inputBuffer = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if ($null -ne $MarkdownText) {
            $inputBuffer.Add($MarkdownText)
        }
    }

    end {
        if ($inputBuffer.Count -eq 0) { return "" }

        # 2. Визначення ширини консолі
        $consoleWidth = 80
        try {
            if ($Host.UI.RawUI.WindowSize.Width -gt 20) {
                $consoleWidth = $Host.UI.RawUI.WindowSize.Width - 2
            }
        } catch {}

        $fullText = $inputBuffer -join "`n"
        $lines = $fullText -split '\r?\n'

        # ВИПРАВЛЕНО ДЛЯ PS 5.1: Екранування символу ESC $e для запобігання помилки індексації $esc[...]
        $e = [char]27
        $reset         = "$($e)[0m"
        $bold          = "$($e)[1m"
        $italic        = "$($e)[3m"
        $underline     = "$($e)[4m"
        $strikethrough = "$($e)[9m"

        # Палітра Dracula (ANSI TrueColor з виправленим синтаксисом $($e)[...)
        $colorH1          = "$($e)[38;2;255;121;198m$bold"
        $colorH2          = "$($e)[38;2;189;147;249m$bold"
        $colorH3          = "$($e)[38;2;139;233;253m$bold"
        $colorH4          = "$($e)[38;2;241;250;140m$bold"
        $colorH5          = "$($e)[38;2;80;250;123m$bold"
        $colorH6          = "$($e)[38;2;98;114;164m$bold"
        $colorCode        = "$($e)[38;2;248;248;242m"
        $colorInlineCode  = "$($e)[38;2;255;184;108m"
        $colorBold        = "$($e)[38;2;255;255;255m$bold"
        $colorItalic      = "$italic$($e)[38;2;248;248;242m"
        $colorBullet      = "$($e)[38;2;241;250;140m"
        $colorBorder      = "$($e)[38;2;98;114;164m"
        $colorQuote       = "$italic$($e)[38;2;98;114;164m"
        $colorLink        = "$($e)[38;2;80;250;123m$underline"
        $colorTableHeader = "$($e)[38;2;139;233;253m$bold"
        $colorSuccess     = "$($e)[38;2;80;250;123m"

        # Допоміжна функція: Очищення ANSI для точного вимірювання довжини
        $getVisualLength = {
            param([string]$text)
            $cleanText = $text -replace "\x1b\[[0-9;]*m", ""
            return $cleanText.Length
        }

        # Допоміжна функція: Рендеринг таблиці
        $renderTable = {
            param([System.Collections.Generic.List[string]]$tableBuffer)
            if ($null -eq $tableBuffer -or $tableBuffer.Count -lt 2) {
                return $tableBuffer
            }

            $parsedRows = [System.Collections.Generic.List[array]]::new()
            foreach ($tLine in $tableBuffer) {
                $rawCells = $tLine.Trim() -replace '^\||\|$', '' -split '\|'
                $cleanCells = foreach ($c in $rawCells) { $c.Trim() }
                $parsedRows.Add($cleanCells)
            }

            $sepIndex = -1
            for ($i = 0; $i -lt $parsedRows.Count; $i++) {
                $isSep = $true
                foreach ($cell in $parsedRows[$i]) {
                    if ($cell -notmatch '^:?-+:?$') {
                        $isSep = $false
                        break
                    }
                }
                if ($isSep -and $parsedRows[$i].Length -gt 0) {
                    $sepIndex = $i
                    break
                }
            }

            if ($sepIndex -eq -1) { return $tableBuffer }

            $maxCols = 0
            foreach ($row in $parsedRows) {
                if ($row.Length -gt $maxCols) { $maxCols = $row.Length }
            }

            $colWidths = New-Object int[] $maxCols
            for ($c = 0; $c -lt $maxCols; $c++) {
                $maxLen = 3
                for ($r = 0; $r -lt $parsedRows.Count; $r++) {
                    if ($r -eq $sepIndex) { continue }
                    if ($c -lt $parsedRows[$r].Length) {
                        $visLen = & $getVisualLength $parsedRows[$r][$c]
                        if ($visLen -gt $maxLen) { $maxLen = $visLen }
                    }
                }
                $colWidths[$c] = $maxLen
            }

            $topParts = foreach ($w in $colWidths) { [string]::new([char]9472, $w + 2) }
            $topBorder = "$colorBorder┌" + ($topParts -join '┬') + "┐$reset"

            $midParts = foreach ($w in $colWidths) { [string]::new([char]9472, $w + 2) }
            $midBorder = "$colorBorder├" + ($midParts -join '┼') + "┤$reset"

            $botParts = foreach ($w in $colWidths) { [string]::new([char]9472, $w + 2) }
            $bottomBorder = "$colorBorder└" + ($botParts -join '┴') + "┘$reset"

            $result = [System.Collections.Generic.List[string]]::new()
            $result.Add($topBorder)

            for ($r = 0; $r -lt $parsedRows.Count; $r++) {
                if ($r -eq $sepIndex) {
                    $result.Add($midBorder)
                    continue
                }

                $rowCells = [System.Collections.Generic.List[string]]::new()
                for ($c = 0; $c -lt $maxCols; $c++) {
                    $cellText = if ($c -lt $parsedRows[$r].Length) { $parsedRows[$r][$c] } else { "" }

                    if ($r -lt $sepIndex) {
                        $cellText = "$colorTableHeader$cellText$reset"
                    }

                    $visLen = & $getVisualLength $cellText
                    $padLen = [Math]::Max(0, $colWidths[$c] - $visLen)
                    $paddedCell = " " + $cellText + ([string]::new(' ', $padLen)) + " "
                    $rowCells.Add($paddedCell)
                }

                $result.Add("$colorBorder│$reset" + ($rowCells -join "$colorBorder│$reset") + "$colorBorder│$reset")
            }

            $result.Add($bottomBorder)
            return $result
        }

        # Допоміжна функція для обробки Inline-форматування (код, посилання, жирний, курсив)
        $applyInlineFormatting = {
            param([string]$text, [string]$lineColor)

            $lineCodeBlocks = [System.Collections.Generic.List[string]]::new()
            $tempLine = $text

            # 1. Маскуємо inline-код
            $codeRegex = [regex]'`([^`\n]+)`'
            $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
                param($m)
                $lineCodeBlocks.Add($m.Groups[1].Value)
                return "%%CODEBLOCK$($lineCodeBlocks.Count - 1)%%"
            }
            $tempLine = $codeRegex.Replace($tempLine, $evaluator)

            $activeReset = if ($lineColor) { "$reset$lineColor" } else { $reset }

            # 2. Зображення
            $tempLine = $tempLine -replace '!\[([^\]]*)\]\(([^)]+)\)', "🖼️  $colorBold`$1$reset$lineColor ($italic`$2$reset$lineColor)"

            # 3. Автопосилання
            $tempLine = $tempLine -replace '<(https?://[^>]+)>', "$colorLink`$1$reset$lineColor"

            # 4. Посилання
            $tempLine = $tempLine -replace '\[([^\]]+)\]\(([^)]+)\)', "$colorLink`$1$reset$lineColor ($italic`$2$reset$lineColor)"

            # 5. Закреслений текст
            $tempLine = $tempLine -replace '~~([^\s~][^~]*?[^\s~]|\S)~~', "$strikethrough`$1$activeReset"

            # 6. Жирний текст
            $tempLine = $tempLine -replace '(?<!\*)\*\*([^\s*][^*]*?[^\s*]|\S)\*\*(?!\*)', "$colorBold`$1$activeReset"
            $tempLine = $tempLine -replace '(?<!_)__([^\s_][^_]*?[^\s_]|\S)__(?!_)', "$colorBold`$1$activeReset"

            # 7. Курсив
            $tempLine = $tempLine -replace '(?<!\*)\*([^\s*][^*]*?[^\s*]|\S)\*(?!\*)', "$colorItalic`$1$activeReset"
            $tempLine = $tempLine -replace '(?<!_)_([^\s_][^_]*?[^\s_]|\S)_(?!_)', "$colorItalic`$1$activeReset"

            # 8. Відновлюємо inline-код
            for ($i = 0; $i -lt $lineCodeBlocks.Count; $i++) {
                $placeholder = "%%CODEBLOCK$i%%"
                $formattedCode = "$colorInlineCode$($lineCodeBlocks[$i])$activeReset"
                $tempLine = $tempLine.Replace($placeholder, $formattedCode)
            }

            return $tempLine
        }

        $inCodeBlock = $false
        $tableBuffer = [System.Collections.Generic.List[string]]::new()
        $processedLines = [System.Collections.Generic.List[string]]::new()

        foreach ($line in $lines) {

            # 1. ОБРОБКА БЛОКІВ КОДУ
            if ($line -match '^(\s*)(`{3,}|~{3,})(.*)$') {
                if ($tableBuffer.Count -gt 0) {
                    $rendered = & $renderTable $tableBuffer
                    foreach ($rLine in $rendered) { $processedLines.Add($rLine) }
                    $tableBuffer.Clear()
                }

                $indent = $Matches[1]
                $lang = $Matches[3].Trim()
                $width = [Math]::Max(10, $consoleWidth - $indent.Length - 2)

                if ($inCodeBlock) {
                    $inCodeBlock = $false
                    $processedLines.Add("$indent$colorBorder└$([string]::new('─', $width))┘$reset")
                } else {
                    $inCodeBlock = $true
                    $headerText = if ($lang) { " CODE: $lang " } else { " CODE " }
                    if ($headerText.Length -gt $width) { $headerText = $headerText.Substring(0, $width) }
                    $paddedHeader = $headerText.PadRight($width, '─')
                    $processedLines.Add("$indent$colorBorder┌$paddedHeader┐$reset")
                }
                continue
            }

            if ($inCodeBlock) {
                $processedLines.Add("$colorCode$line$reset")
                continue
            }

            # 2. ДЕТЕКЦІЯ ТАБЛИЦЬ
            if ($line -match '\|') {
                # Застосовуємо форматування всередині комірок
                $tempTableLine = & $applyInlineFormatting $line ""
                $tableBuffer.Add($tempTableLine)
                continue
            } else {
                if ($tableBuffer.Count -gt 0) {
                    $rendered = & $renderTable $tableBuffer
                    foreach ($rLine in $rendered) { $processedLines.Add($rLine) }
                    $tableBuffer.Clear()
                }
            }

            # 3. БЛОЧНІ ЕЛЕМЕНТИ (Заголовки, Списки, Чекбокси)
            $lineColor = ""
            $prefix = ""
            $content = $line

            if ($line -match '^#\s+(.*)$')          { $lineColor = $colorH1; $prefix = "# "; $content = $Matches[1] }
            elseif ($line -match '^##\s+(.*)$')     { $lineColor = $colorH2; $prefix = "## "; $content = $Matches[1] }
            elseif ($line -match '^###\s+(.*)$')    { $lineColor = $colorH3; $prefix = "### "; $content = $Matches[1] }
            elseif ($line -match '^####\s+(.*)$')   { $lineColor = $colorH4; $prefix = "#### "; $content = $Matches[1] }
            elseif ($line -match '^#####\s+(.*)$')  { $lineColor = $colorH5; $prefix = "##### "; $content = $Matches[1] }
            elseif ($line -match '^######+\s+(.*)$'){ $lineColor = $colorH6; $prefix = "###### "; $content = $Matches[1] }

            # Чекбокси Невиконані [-]
            elseif ($line -match '^(\s*)[*\-]\s+\[\s\]\s+(.*)$') {
                $prefix = "$($Matches[1])$colorBorder☐$reset "
                $content = $Matches[2]
            }
            # Чекбокси Виконані [x]
            elseif ($line -match '^(\s*)[*\-]\s+\[[xX]\]\s+(.*)$') {
                $prefix = "$($Matches[1])$colorSuccess☑$reset "
                $content = $Matches[2]
            }
            # Марковані списки
            elseif ($line -match '^(\s*)[*\-]\s+(.*)$') {
                $prefix = "$($Matches[1])$colorBullet•$reset "
                $content = $Matches[2]
            }
            # Нумеровані списки
            elseif ($line -match '^(\s*)(\d+)\.\s+(.*)$') {
                $prefix = "$($Matches[1])$colorBullet$($Matches[2]).$reset "
                $content = $Matches[3]
            }
            # Цитати
            elseif ($line -match '^(\s*)>\s+(.*)$') {
                $lineColor = $colorQuote
                $prefix = "$($Matches[1])$colorBorder│$reset "
                $content = $Matches[2]
            }
            # Горизонтальна лінія
            elseif ($line -match '^[-*_]{3,}$') {
                $lineWidth = [Math]::Max(10, $consoleWidth)
                $processedLines.Add("$colorBorder$([string]::new('─', $lineWidth))$reset")
                continue
            }

            # 4. ОБРОБКА ВНУТРІШНЬОГО ТЕКСТУ ТА ЗБИРАННЯ РЯДКА
            $formattedContent = & $applyInlineFormatting $content $lineColor

            if ($lineColor) {
                $finalLine = "$lineColor$prefix$formattedContent$reset"
            } else {
                $finalLine = "$prefix$formattedContent"
            }

            $processedLines.Add($finalLine)
        }

        # Очищення остаточного буфера таблиць
        if ($tableBuffer.Count -gt 0) {
            $rendered = & $renderTable $tableBuffer
            foreach ($rLine in $rendered) { $processedLines.Add($rLine) }
            $tableBuffer.Clear()
        }

        # Відновлюємо кодування
        if ($null -ne $script:originalEncoding) {
            [Console]::OutputEncoding = $script:originalEncoding
        }

        $processedLines -join [char]10
    }
}

# ===================================================================
# ФУНКЦІЇ РОБОТИ З БУФЕРОМ ОБМІНУ 
# ===================================================================

function Get-ClipboardText {
    try {
        # -ErrorAction Stop перехоплює помилки блокування Win32 API (CLIPBRD_E_CANT_OPEN)
        $text = Get-Clipboard -Raw -ErrorAction Stop

        if (-not [string]::IsNullOrWhiteSpace($text)) {
            return $text
        }
    }
    catch {
        # Помилка виникає, якщо буфер містить картинку/файли або заблокований
        Write-Muted "ℹ️ Не вдалося зчитати текст з буфера обміну (буфер порожній або заблокований)."
    }
    return $null
}

function Set-ClipboardText ([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    try {
        Set-Clipboard -Value $text -ErrorAction Stop
        return $true
    }
    catch {
        Write-ErrorMsg "Помилка запису в буфер обміну: $_"
        return $false
    }
}

# ===================================================================
# РОЗУМНА СИСТЕМА ВВОДУ 
# ===================================================================

function Read-SmartInput {
    [CmdletBinding()]
    param (
        [float]$TimeoutSeconds = 2.5 # Затримка відправки у секундах
    )

    $lines = [System.Collections.Generic.List[string]]::new()

    while ($true) {
        # Візуальний маркер продовжуваного вводу
        if ($lines.Count -gt 0) {
            Write-Host "$($script:UI.Comment)  │ $($script:UI.Reset)" -NoNewline
        }

        # 1. Нативне читання рядка
        $line = Read-Host
        $trimmed = $line.Trim()

        # 2. Обробка системних команд на першому рядку (виконуються негайно)
        if ($lines.Count -eq 0 -and $trimmed.StartsWith("/") -and $trimmed -ne "/p") {
            return $trimmed
        }

        # 3. Обробка команди вставки /p з повним виведенням вмісту на екран
        if ($trimmed -eq "/p") {
            $clip = Get-ClipboardText
            if (-not [string]::IsNullOrWhiteSpace($clip)) {
                # Нормалізація та розбиття тексту з буфера на рядки
                $clipLines = $clip -split '\r?\n'

                # Візуальний друк вставленого тексту для перевірки користувачем
                foreach ($cL in $clipLines) {
                    Write-Host "$($script:UI.Comment)  │ $($script:UI.Green)$cL$($script:UI.Reset)"
                    $lines.Add($cL)
                }

                Write-Success "📋 Успішно вставлено та додано $($clipLines.Count) рядків з буфера обміну."
            } else {
                Write-Warn "⚠️ Буфер обміну порожній або не містить текстових даних."
            }
        }

        # 4. Накопичення звичайного тексту
        elseif (-not [string]::IsNullOrWhiteSpace($line) -or $lines.Count -gt 0) {
            $lines.Add($line)
        }

        # Якщо нічого не введено на першому рядку (просто натиснули Enter) — скасовуємо
        if ($lines.Count -eq 0) { return "" }

        # 5. Цикл очікування (Debounce Timer)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $keyPressed = $false

        while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
            if ([Console]::KeyAvailable) {
                $keyPressed = $true
                break # Натиснуто клавішу -> йдемо на наступне коло Read-Host
            }

            # Анімація зворотного відліку
            $rem = [Math]::Ceiling($TimeoutSeconds - $sw.Elapsed.TotalSeconds)
            Write-Host "`r$($script:UI.Comment)  ⏳ Відправка через ${rem}с... (натисніть клавішу для продовження)$($script:UI.Reset)" -NoNewline

            Start-Sleep -Milliseconds 50
        }

        # Очищення рядка таймера
        Write-Host "`r                                                                       `r" -NoNewline

        # 6. Якщо таймер вийшов без нових натискань — повертаємо накопичений текст
        if (-not $keyPressed) {
            return ($lines -join "`n").Trim()
        }
    }
}


# ===================================================================
# --- Основний робочий процес CLI ---
# ===================================================================
# Виклик функції стартового екрану при запуску
Clear-Host

Reset-Session
Show-WelcomeScreen

$currentKey = Get-ApiKey

while ($true) {
    # Індикатор збереження діалогу
	# $indicator = if ($script:CurrentSessionFile) { " 💾" } else { "" } 
    $indicator = switch ($true) {
        # Файл є + Є незбережені зміни
        ($script:CurrentSessionFile -and $script:IsModified) { " 💾✏️" ; break }
        # Файлу немає + Є незбережені зміни (новий діалог)
        (-not $script:CurrentSessionFile -and $script:IsModified) { " 📝" ; break }
        # Файл є + Змін немає (повністю збережено)
        ($script:CurrentSessionFile -and -not $script:IsModified) { " 💾" ; break }
        # Чистий стан
        Default { "" }
    }    

	Write-Host "$($script:UI.Green)$($script:UI.Bold)`n🧑🏻 Користувач$indicator : $($script:UI.Reset)" -NoNewline

    # Отримання вводу за допомогою нашого розумного движка (з затримкою 2.5 сек)
    $userInput = Read-SmartInput -TimeoutSeconds 5

    if ([string]::IsNullOrWhiteSpace($userInput)) { continue }
    $userInput = $userInput.Trim()

    # --- Обробка системних команд ---
    
    if ($userInput -eq "/c") {
        $lastAiMessage = $script:ChatHistory | Where-Object { $_.role -eq "assistant" } | Select-Object -Last 1

        if ($lastAiMessage -and -not [string]::IsNullOrWhiteSpace($lastAiMessage.content)) {
            if (Set-ClipboardText -text $lastAiMessage.content) {
                Write-Success "📋️ Останню відповідь AI успішно скопійовано в буфер обміну!"
            }
        } else {
            Write-Warn "Немає відповідей AI у поточному діалозі для копіювання."
        }
        continue
    }    
    
    if ($userInput -eq "/exit") {
        # Перевіряємо, чи є повідомлення і чи були вони змінені/додані з моменту останнього збереження
        if ($script:ChatHistory.Count -gt 1 -and $script:IsModified) {
            Write-Warn "Є незбережені зміни. Зберегти перед виходом? (y/n): " -NoNewline
            $confirm = Read-Host
            if ($confirm -eq "y" -or $confirm -eq "yes") {
                Save-ChatHistory
            }
        }
        Write-Info "Дякуємо за використання! До побачення! 👋"
        Start-Sleep -Seconds 3
        break
    }

    if ($userInput -eq "/new") {
        # Перевіряємо, чи є повідомлення і чи були вони змінені/додані з моменту останнього збереження
        if ($script:ChatHistory.Count -gt 1 -and $script:IsModified) {
            Write-Warn "Є незбережені зміни. Зберегти перед виходом? (y/n): " -NoNewline
            $confirm = Read-Host
            if ($confirm -eq "y" -or $confirm -eq "yes") {
                Save-ChatHistory
            }
        }
        Clear-Host
        Reset-Session
        Show-WelcomeScreen

        Write-Warn "🧹 Історію сесії очищено! Системний промпт скинуто до стандартного."
        continue
    }

    if ($userInput -eq "/key") {
        try {
            # Перезапитуємо та перезаписуємо ключ у пам'яті та у файлі
            $currentKey = Get-ApiKey -Force
            Write-Success "Ключ успішно оновлено для поточної сесії!"
        } catch {
            Write-Warn "Зміну ключа скасовано."
        }
        continue
    }

    if ($userInput -eq "/save") {
        Save-ChatHistory
        continue
    }

    if ($userInput -eq "/load") {
        Reset-Session
        Load-ChatHistory
        Refresh-ChatHistory
        continue
    }

    if ($userInput -eq "/export") {
        Export-ToMarkdown
        continue
    }

	if ($userInput -eq "/refresh") {
        Clear-Host
        Refresh-ChatHistory
        continue
    }

    if ($userInput -eq "/help" -or $userInput -eq "/?") {
        Show-WelcomeScreen
        continue
    }

    # Виявлення помилок при введенні команд, що починаються з "/"
    if ($userInput.StartsWith("/")) {
        Write-ErrorMsg "Невідома команда: $userInput"
        Write-Warn "Доступні команди: /new, /save, /load, /export, /refresh, /c, /p, /key, /help, /exit"
        continue
    }

    # --- Стандартний запит до LLM ---

    # 1. Додавання повідомлення користувача
    $script:ChatHistory += @{ role = "user"; content = $userInput }
	$script:IsModified = $true 

    # 2. Серіалізація поточного контексту
    $postBody = @{
        model = $ModelName
        messages = $script:ChatHistory
        temperature = 0.7
    } | ConvertTo-Json -Depth 100

    $headers = @{
        "Authorization" = "Bearer $currentKey"
        "Content-Type"  = "application/json; charset=utf-8"
    }

    # Індикатор думок:
    Write-Warn "🤔 PS-COPILOT генерує відповідь..." -NoNewline
    
    try {
        # 3. POST запит до Gemini (інтерфейс OpenAI API)
        $response = Invoke-RestMethod -Uri $ApiUrl -Method Post -Headers $headers -Body $postBody -ContentType "application/json; charset=utf-8"

        # Очищення індикатора очікування
        Write-Host "`r                                        " -NoNewline
        Write-Host "`r" -NoNewline

		# Отримуємо сирий некоректно декодований текст
		$rawReply = $response.choices[0].message.content

		# Виправлення кодування: повертаємо в байти Latin-1 та декодуємо в UTF-8
        # Фікс кодування потрібен ТІЛЬКИ для PowerShell 5.1
        if ($PSVersionTable.PSVersion.Major -le 5) {
            $replyBytes = [System.Text.Encoding]::GetEncoding("iso-8859-1").GetBytes($rawReply)
            $reply = [System.Text.Encoding]::UTF8.GetString($replyBytes)
        } else {
            $reply = $rawReply
        }
        
        # Вивід відповіді AI з форматуванням
        Write-Host "$($script:UI.Cyan)$($script:UI.Bold)🤖 PS-COPILOT [$ModelName] :$($script:UI.Reset)"
		$ansiReply = Convert-MarkdownToAnsi -MarkdownText $reply
		# $ansiReply = $reply
		Write-Host $ansiReply

        # 4. Додавання відповіді AI до контексту
        $script:ChatHistory += @{ role = "assistant"; content = $reply }
		$script:IsModified = $true
    }
    catch {
        Write-Host "`r" -NoNewline
        Write-ErrorMsg "Помилка запиту до API: $_"
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errText = $reader.ReadToEnd()
                Write-Host "$($script:UI.Red)Повні деталі помилки від сервера:$($script:UI.Reset)"
                Write-Host "$($script:UI.Red)$errText$($script:UI.Reset)"
            } catch {}
        }
        # Відкат останнього вводу користувача, якщо запит не вдався, для збереження цілісності сесії
        if ($script:ChatHistory[-1].role -eq "user") {
            $script:ChatHistory = $script:ChatHistory[0..($script:ChatHistory.Count-2)]
            Write-Warn "Останній ввід видалено з пам'яті для збереження цілісності сесії."
        }
    }
}
