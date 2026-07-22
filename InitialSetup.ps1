<#
.SYNOPSIS
    Бэкап, настройка IP, ввод в домен, прокси, выбор OU, авто-поиск свободного IP.
    Версия 2.0 (с улучшениями)
#>

# Проверка прав администратора
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ОШИБКА: Запустите скрипт от имени Администратора!" -ForegroundColor Red
    pause
    exit 1
}

# === 1. БЭКАП ===
$backupDir = "$env:USERPROFILE\Desktop\NetworkBackup"
if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

# Получаем все активные физические адаптеры для выбора
$adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' }
if (-not $adapters) {
    Write-Host "Не найдено активных физических адаптеров!" -ForegroundColor Red
    pause
    exit 1
}

# === 2. ФУНКЦИИ ВАЛИДАЦИИ (НОВОЕ) ===
function Test-ValidIP {
    param([string]$IP)
    $regex = '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return ($IP -match $regex)
}

function Test-ValidPrefix {
    param([string]$Prefix)
    return ($Prefix -match '^([0-9]|[1-2][0-9]|3[0-2])$' -and [int]$Prefix -ge 1 -and [int]$Prefix -le 32)
}

function Test-ValidComputerName {
    param([string]$Name)
    $regex = '^[a-zA-Z0-9-]{1,15}$'
    return ($Name -match $regex)
}

# === 3. ФУНКЦИЯ ПОИСКА СВОБОДНОГО IP (НОВОЕ) ===
function Find-FreeIP {
    param($Gateway, $PrefixLength)
    # Вычисляем диапазон по шлюзу и маске
    $gwBytes = [IPAddress]::Parse($Gateway).GetAddressBytes()
    $maskBytes = switch ($PrefixLength) {
        24 { @(255,255,255,0) }
        16 { @(255,255,0,0) }
        8  { @(255,0,0,0) }
        default { # для упрощения только /24,/16,/8, но можно расширить
            Write-Host "Автопоиск поддерживается только для /8, /16, /24" -ForegroundColor Yellow
            return $null
        }
    }
    $network = @()
    $broadcast = @()
    for ($i=0; $i -lt 4; $i++) {
        $network += $gwBytes[$i] -band $maskBytes[$i]
        $broadcast += $gwBytes[$i] -bor (-bnot $maskBytes[$i] -band 0xFF)
    }
    $first = $network.Clone()
    $first[3] += 1
    $last = $broadcast.Clone()
    $last[3] -= 1

    # Перебираем адреса от first до last и пингуем
    for ($i = $first[3]; $i -le $last[3]; $i++) {
        $testIP = "$($first[0]).$($first[1]).$($first[2]).$i"
        if (-not (Test-Connection -ComputerName $testIP -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            return $testIP
        }
    }
    return $null
}

# === 4. ФУНКЦИЯ ПРОВЕРКИ ДОСТУПНОСТИ DC (НОВОЕ) ===
function Test-DomainController {
    param($Domain)
    try {
        $dc = (Resolve-DnsName -Name $Domain -Type A -ErrorAction Stop).IPAddress
        if ($dc) {
            return (Test-Connection -ComputerName $dc -Count 1 -Quiet)
        }
    } catch {
        return $false
    }
    return $false
}

# === 5. WPF ОКНО (расширенное) ===
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Первоначальная настройка ПК v2.0" Height="730" Width="600"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="12">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Группа "Сетевой адаптер" (НОВОЕ) -->
        <GroupBox Header="Сетевой адаптер" Grid.Row="0" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Grid.Column="0" Content="Выберите адаптер:" VerticalAlignment="Center"/>
                <ComboBox x:Name="cmbAdapters" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
                <Button x:Name="btnRefreshAdapters" Grid.Column="2" Content="Обновить" Margin="10,2,0,2" Width="80"/>
            </Grid>
        </GroupBox>

        <!-- Группа IP -->
        <GroupBox Header="Параметры IP" Grid.Row="1" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                </Grid.RowDefinitions>

                <Label Grid.Row="0" Grid.Column="0" Content="Новый IP-адрес:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtIP" Grid.Row="0" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
                <Button x:Name="btnFindFreeIP" Grid.Row="0" Grid.Column="2" Content="Найти свободный" Margin="5,2,0,2" Width="120" ToolTip="Ищет свободный IP в подсети шлюза"/>

                <Label Grid.Row="1" Grid.Column="0" Content="Новая маска (префикс):" VerticalAlignment="Center"/>
                <TextBox x:Name="txtPrefix" Grid.Row="1" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center" Text="24" Width="60" HorizontalAlignment="Left"/>

                <Label Grid.Row="2" Grid.Column="0" Content="Новый шлюз:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtGateway" Grid.Row="2" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                <Label Grid.Row="3" Grid.Column="0" Content="Новый DNS (предпоч.):" VerticalAlignment="Center"/>
                <TextBox x:Name="txtDNS1" Grid.Row="3" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                <Label Grid.Row="4" Grid.Column="0" Content="Новый DNS (запасной):" VerticalAlignment="Center"/>
                <TextBox x:Name="txtDNS2" Grid.Row="4" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
            </Grid>
        </GroupBox>

        <!-- Группа ПК/Домен (добавлено OU) -->
        <GroupBox Header="Компьютер и домен" Grid.Row="2" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                </Grid.RowDefinitions>

                <Label Grid.Row="0" Grid.Column="0" Content="Новое имя ПК:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtComputerName" Grid.Row="0" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                <Label Grid.Row="1" Grid.Column="0" Content="Новый домен:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtDomain" Grid.Row="1" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                <Label Grid.Row="2" Grid.Column="0" Content="OU (подразделение):" VerticalAlignment="Center" ToolTip="Оставьте пустым для размещения в контейнере по умолчанию"/>
                <TextBox x:Name="txtOU" Grid.Row="2" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
            </Grid>
        </GroupBox>

        <!-- Группа Прокси (НОВОЕ) -->
        <GroupBox Header="Настройка прокси (опционально)" Grid.Row="3" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                </Grid.RowDefinitions>

                <CheckBox x:Name="chkProxy" Grid.Row="0" Grid.ColumnSpan="3" Content="Настроить системный прокси" VerticalAlignment="Center" Margin="0,2,0,2"/>
                <Label Grid.Row="1" Grid.Column="0" Content="Адрес прокси:" VerticalAlignment="Center" IsEnabled="False"/>
                <TextBox x:Name="txtProxyAddress" Grid.Row="1" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center" IsEnabled="False"/>
                <Label Grid.Row="1" Grid.Column="2" Content="Порт:" VerticalAlignment="Center" IsEnabled="False" Margin="5,0,0,0"/>
                <TextBox x:Name="txtProxyPort" Grid.Row="1" Grid.Column="2" Width="60" Margin="45,2,0,2" VerticalAlignment="Center" IsEnabled="False"/>
            </Grid>
        </GroupBox>

        <!-- Прогресс-бар (НОВОЕ) -->
        <ProgressBar x:Name="progressBar" Grid.Row="4" Height="20" Margin="0,5,0,10" IsIndeterminate="False" Minimum="0" Maximum="100" Value="0"/>

        <!-- Кнопка OK -->
        <StackPanel Grid.Row="5" HorizontalAlignment="Center" Margin="0,10,0,0" Orientation="Horizontal">
            <Button x:Name="btnOK" Content="OK" Width="100" Height="30" IsDefault="True" Margin="0,0,10,0">
                <Button.Background>
                    <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                        <GradientStop Color="Black"/>
                        <GradientStop Color="#FFBFBBBB" Offset="1"/>
                    </LinearGradientBrush>
                </Button.Background>
            </Button>
            <Button x:Name="btnCancel" Content="Отмена" Width="100" Height="30" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Создание окна из XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Получение ссылок на элементы
$cmbAdapters = $window.FindName("cmbAdapters")
$btnRefreshAdapters = $window.FindName("btnRefreshAdapters")
$txtIP = $window.FindName("txtIP")
$txtPrefix = $window.FindName("txtPrefix")
$txtGateway = $window.FindName("txtGateway")
$txtDNS1 = $window.FindName("txtDNS1")
$txtDNS2 = $window.FindName("txtDNS2")
$txtComputerName = $window.FindName("txtComputerName")
$txtDomain = $window.FindName("txtDomain")
$txtOU = $window.FindName("txtOU")
$chkProxy = $window.FindName("chkProxy")
$txtProxyAddress = $window.FindName("txtProxyAddress")
$txtProxyPort = $window.FindName("txtProxyPort")
$btnFindFreeIP = $window.FindName("btnFindFreeIP")
$progressBar = $window.FindName("progressBar")
$btnOK = $window.FindName("btnOK")
$btnCancel = $window.FindName("btnCancel")

# Заполнение списка адаптеров
function UpdateAdapters {
    $cmbAdapters.Items.Clear()
    $adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' }
    foreach ($ad in $adapters) {
        $cmbAdapters.Items.Add("$($ad.Name) (Index $($ad.InterfaceIndex))")
    }
    if ($cmbAdapters.Items.Count -gt 0) { $cmbAdapters.SelectedIndex = 0 }
}
UpdateAdapters

# Обновление адаптеров по кнопке
$btnRefreshAdapters.Add_Click({ UpdateAdapters })

# Подстановка текущего имени компьютера
$txtComputerName.Text = $env:COMPUTERNAME

# Включение/отключение полей прокси
$chkProxy.Add_Checked({
    $txtProxyAddress.IsEnabled = $true
    $txtProxyPort.IsEnabled = $true
})
$chkProxy.Add_Unchecked({
    $txtProxyAddress.IsEnabled = $false
    $txtProxyPort.IsEnabled = $false
})

# Кнопка "Найти свободный IP"
$btnFindFreeIP.Add_Click({
    $gw = $txtGateway.Text.Trim()
    $prefix = $txtPrefix.Text.Trim()
    if (-not (Test-ValidIP $gw)) {
        [System.Windows.MessageBox]::Show("Введите корректный IP шлюза перед поиском.", "Ошибка", "OK", "Error")
        return
    }
    if (-not (Test-ValidPrefix $prefix)) {
        [System.Windows.MessageBox]::Show("Введите корректный префикс (1-32).", "Ошибка", "OK", "Error")
        return
    }
    $freeIP = Find-FreeIP -Gateway $gw -PrefixLength ([int]$prefix)
    if ($freeIP) {
        $txtIP.Text = $freeIP
        [System.Windows.MessageBox]::Show("Найден свободный IP: $freeIP", "Успех", "OK", "Information")
    } else {
        [System.Windows.MessageBox]::Show("Не удалось найти свободный IP в диапазоне.", "Предупреждение", "OK", "Warning")
    }
})

# Обработчик кнопки OK с валидацией
$btnOK.Add_Click({
    # Валидация
    $ip = $txtIP.Text.Trim()
    $prefix = $txtPrefix.Text.Trim()
    $gw = $txtGateway.Text.Trim()
    $dns1 = $txtDNS1.Text.Trim()
    $dns2 = $txtDNS2.Text.Trim()
    $compName = $txtComputerName.Text.Trim()
    $domain = $txtDomain.Text.Trim()
    $ou = $txtOU.Text.Trim()

    if (-not (Test-ValidIP $ip)) {
        [System.Windows.MessageBox]::Show("Некорректный IP-адрес.", "Ошибка", "OK", "Error")
        return
    }
    if (-not (Test-ValidPrefix $prefix)) {
        [System.Windows.MessageBox]::Show("Некорректный префикс (1-32).", "Ошибка", "OK", "Error")
        return
    }
    if (-not (Test-ValidIP $gw)) {
        [System.Windows.MessageBox]::Show("Некорректный шлюз.", "Ошибка", "OK", "Error")
        return
    }
    if (-not (Test-ValidIP $dns1)) {
        [System.Windows.MessageBox]::Show("Некорректный DNS1.", "Ошибка", "OK", "Error")
        return
    }
    if ($dns2 -and -not (Test-ValidIP $dns2)) {
        [System.Windows.MessageBox]::Show("Некорректный DNS2 (если не используется, оставьте пустым).", "Ошибка", "OK", "Error")
        return
    }
    if (-not (Test-ValidComputerName $compName)) {
        [System.Windows.MessageBox]::Show("Некорректное имя компьютера (до 15 символов, только буквы, цифры, дефис).", "Ошибка", "OK", "Error")
        return
    }
    if (-not $domain) {
        [System.Windows.MessageBox]::Show("Поле 'Домен' обязательно.", "Ошибка", "OK", "Error")
        return
    }

    # Проверка доступности DC (НОВОЕ)
    if (-not (Test-DomainController -Domain $domain)) {
        $res = [System.Windows.MessageBox]::Show("Контроллер домена $domain не доступен. Продолжить?", "Предупреждение", "YesNo", "Warning")
        if ($res -eq 'No') { return }
    }

    # Проверка прокси
    if ($chkProxy.IsChecked) {
        $proxyAddr = $txtProxyAddress.Text.Trim()
        $proxyPort = $txtProxyPort.Text.Trim()
        if (-not $proxyAddr -or -not $proxyPort) {
            [System.Windows.MessageBox]::Show("Для настройки прокси заполните адрес и порт.", "Ошибка", "OK", "Error")
            return
        }
        if (-not ($proxyPort -match '^\d+$' -and [int]$proxyPort -gt 0 -and [int]$proxyPort -le 65535)) {
            [System.Windows.MessageBox]::Show("Некорректный порт прокси.", "Ошибка", "OK", "Error")
            return
        }
    }

    # Если всё ок, закрываем окно с успехом
    $window.DialogResult = $true
    $window.Close()
})

$btnCancel.Add_Click({
    $window.DialogResult = $false
    $window.Close()
})

# Показываем окно
$result = $window.ShowDialog()

if ($result -ne $true) {
    Write-Host "Настройка отменена пользователем." -ForegroundColor Yellow
    exit
}

# Извлекаем значения (уже проверены)
$selectedAdapterString = $cmbAdapters.SelectedItem
$adapterIndex = [int]($selectedAdapterString -split '\(Index ')[1] -replace '\)',''
$newIP = $txtIP.Text.Trim()
$prefixLength = [int]$txtPrefix.Text.Trim()
$gateway = $txtGateway.Text.Trim()
$dns1 = $txtDNS1.Text.Trim()
$dns2 = $txtDNS2.Text.Trim()
$newName = $txtComputerName.Text.Trim()
$domain = $txtDomain.Text.Trim()
$ou = $txtOU.Text.Trim()
$proxyEnabled = $chkProxy.IsChecked
$proxyAddress = $txtProxyAddress.Text.Trim()
$proxyPort = $txtProxyPort.Text.Trim()

# Логируем в консоль
Write-Host "Выбран адаптер: $selectedAdapterString" -ForegroundColor Cyan
Write-Host "IP: $newIP/$prefixLength, GW: $gateway, DNS: $dns1, $dns2"
Write-Host "Имя: $newName, Домен: $domain, OU: $ou"
if ($proxyEnabled) { Write-Host "Прокси: $proxyAddress`:$proxyPort" }

# === 6. ПРИМЕНЕНИЕ НАСТРОЕК ===

# Создаём бэкап текущих настроек выбранного адаптера
$currentIP = Get-NetIPAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -ne 'WellKnown' }
$currentDNS = Get-DnsClientServerAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4
$currentGateway = Get-NetRoute -InterfaceIndex $adapterIndex -DestinationPrefix '0.0.0.0/0' | Select-Object -ExpandProperty NextHop

$backupFile = "$backupDir\NetworkBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
@"
=== РЕЗЕРВНАЯ КОПИЯ СЕТЕВЫХ НАСТРОЕК ===
Дата создания: $(Get-Date)
Адаптер: $selectedAdapterString
IPv4 адрес: $($currentIP.IPAddress)/$($currentIP.PrefixLength)
Шлюз: $currentGateway
DNS серверы: $($currentDNS.ServerAddresses -join ', ')
"@ | Out-File -FilePath $backupFile -Encoding UTF8
Write-Host "Бэкап сохранён в: $backupFile" -ForegroundColor Green

# --- Настройка IP ---
Write-Host "`nНазначение статического IP..." -ForegroundColor Yellow
$progressBar.Value = 10
Get-NetIPAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4 |
    Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
    Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
New-NetIPAddress -InterfaceIndex $adapterIndex -IPAddress $newIP -PrefixLength $prefixLength -DefaultGateway $gateway | Out-Null
Set-DnsClientServerAddress -InterfaceIndex $adapterIndex -ServerAddresses ($dns1, $dns2) | Out-Null
Write-Host "IP $newIP/$prefixLength, шлюз $gateway, DNS $dns1 $dns2 применены." -ForegroundColor Green
$progressBar.Value = 30

# --- Настройка прокси (НОВОЕ) ---
if ($proxyEnabled) {
    Write-Host "Настройка системного прокси..." -ForegroundColor Yellow
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 1
    Set-ItemProperty -Path $regPath -Name ProxyServer -Value "$proxyAddress`:$proxyPort"
    # Также для WinHTTP
    netsh winhttp set proxy $proxyAddress`:$proxyPort | Out-Null
    Write-Host "Прокси настроен: $proxyAddress`:$proxyPort" -ForegroundColor Green
}
$progressBar.Value = 40

# --- Переименование ---
$currentName = $env:COMPUTERNAME
if ($newName -ne $currentName) {
    Write-Host "Переименование компьютера в '$newName'..." -ForegroundColor Yellow
    Rename-Computer -NewName $newName -Force
    Write-Host "Имя изменено. Вступит в силу после перезагрузки." -ForegroundColor Green
}
$progressBar.Value = 50

# --- Ввод в домен с OU ---
Write-Host "Добавление в домен '$domain'..." -ForegroundColor Yellow
$cred = Get-Credential -Message "Введите учётные данные администратора домена (DOMAIN\Admin или admin@domain.local)"
if (-not $cred) {
    Write-Host "Учётные данные не введены. Выход." -ForegroundColor Red
    pause
    exit 1
}
try {
    if ($ou) {
        Add-Computer -DomainName $domain -Credential $cred -OUPath $ou -Force -ErrorAction Stop
    } else {
        Add-Computer -DomainName $domain -Credential $cred -Force -ErrorAction Stop
    }
    Write-Host "Компьютер успешно добавлен в домен $domain." -ForegroundColor Green
} catch {
    Write-Host "ОШИБКА при добавлении в домен: $_" -ForegroundColor Red
    Write-Host "Сетевые настройки сохранены. Бэкап доступен в $backupFile" -ForegroundColor Yellow
    pause
    exit 1
}
$progressBar.Value = 80

# === 7. ФИНАЛЬНЫЙ ЛОГ ===
$logFile = "$backupDir\SetupLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
@"
=================================
Лог начальной настройки ПК
Дата: $(Get-Date)
Адаптер: $selectedAdapterString
Компьютер: $currentName -> $newName
Домен: $domain
OU: $ou
Новый IP: $newIP/$prefixLength
Шлюз: $gateway
DNS1: $dns1
DNS2: $dns2
Прокси: $(if ($proxyEnabled) { "$proxyAddress`:$proxyPort" } else { "Не настроен" })
Статус: Успешно
=================================
"@ | Out-File $logFile -Encoding UTF8
$progressBar.Value = 100

Write-Host "`nЛог операций записан в $logFile" -ForegroundColor Cyan

# === 8. ПЕРЕЗАГРУЗКА ===
$restart = Read-Host "`nДля завершения требуется перезагрузка. Перезагрузить сейчас? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Write-Host "Перезагрузка через 10 секунд..." -ForegroundColor Magenta
    shutdown /r /t 10 /c "Завершение настройки ПК"
} else {
    Write-Host "Перезагрузка отложена. Не забудьте перезагрузить компьютер вручную!" -ForegroundColor Yellow
}

pause