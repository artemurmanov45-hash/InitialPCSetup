<#
.SYNOPSIS
    Бэкап, настройка IP, ввод в домен, прокси, выбор OU.
    Версия 2.3.0 (маска в формате 255.255.255.0, обновлены подписи DNS)
#>

# === ПРОВЕРКА ПРАВ АДМИНИСТРАТОРА И АВТО-ПЕРЕЗАПУСК ===
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "Запрашиваем права администратора..." -ForegroundColor Yellow
    Start-Sleep -Milliseconds 500
    
    $scriptPath = "`"" + $MyInvocation.MyCommand.Path + "`""
    $arguments = "-NoProfile -ExecutionPolicy RemoteSigned -File $scriptPath"
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    
    exit
}
# --- Конец проверки ---

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

# === 2. ФУНКЦИИ ВАЛИДАЦИИ ===
function Test-ValidIP {
    param([string]$IP)
    $regex = '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return ($IP -match $regex)
}

# НОВАЯ ФУНКЦИЯ: проверка корректности маски подсети (255.255.255.0 и т.п.)
function Test-ValidSubnetMask {
    param([string]$Mask)
    # Проверяем формат IP
    if (-not (Test-ValidIP $Mask)) { return $false }
    # Разбиваем на октеты
    $octets = $Mask.Split('.')
    $binary = ''
    foreach ($octet in $octets) {
        $binary += [Convert]::ToString([int]$octet, 2).PadLeft(8, '0')
    }
    # Маска должна быть вида: сначала единицы, потом нули (непрерывно)
    if ($binary -match '^1+0+$') {
        return $true
    } else {
        return $false
    }
}

# Функция преобразования маски из точечной нотации в длину префикса
function Convert-MaskToPrefixLength {
    param([string]$Mask)
    $octets = $Mask.Split('.')
    $binary = ''
    foreach ($octet in $octets) {
        $binary += [Convert]::ToString([int]$octet, 2).PadLeft(8, '0')
    }
    return ($binary -replace '0','').Length
}

function Test-ValidComputerName {
    param([string]$Name)
    $regex = '^[a-zA-Z0-9-]{1,15}$'
    return ($Name -match $regex)
}

# === 3. ФУНКЦИЯ ПРОВЕРКИ ДОСТУПНОСТИ DC ===
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

# === 4. WPF ОКНО (ДВУХШАГОВЫЙ МАСТЕР) ===
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Мастер первоначальной настройки ПК v2.3.0" Height="500" Width="580"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="12">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Заголовок шага -->
        <TextBlock x:Name="lblStep" Grid.Row="0" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>

        <!-- Контейнер для шагов -->
        <Grid Grid.Row="1">
            <!-- ШАГ 1: Сетевые параметры -->
            <StackPanel x:Name="step1" Visibility="Visible">
                <GroupBox Header="Сетевой адаптер" Padding="10" Margin="0,0,0,10">
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

                <GroupBox Header="Параметры IP" Padding="10" Margin="0,0,0,10">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
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

                        <!-- Изменено: поле для маски в формате 255.255.255.0 -->
                        <Label Grid.Row="1" Grid.Column="0" Content="Новая маска подсети:" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtMask" Grid.Row="1" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center" Text="255.255.255.0" Width="120" HorizontalAlignment="Left"/>

                        <Label Grid.Row="2" Grid.Column="0" Content="Новый шлюз:" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtGateway" Grid.Row="2" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                        <!-- Изменены подписи DNS -->
                        <Label Grid.Row="3" Grid.Column="0" Content="Предпочитаемый DNS-сервер:" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtDNS1" Grid.Row="3" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>

                        <Label Grid.Row="4" Grid.Column="0" Content="Альтернативный DNS-сервер:" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtDNS2" Grid.Row="4" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
                    </Grid>
                </GroupBox>
            </StackPanel>

            <!-- ШАГ 2: Компьютер, домен, прокси -->
            <StackPanel x:Name="step2" Visibility="Collapsed">
                <GroupBox Header="Компьютер и домен" Padding="10" Margin="0,0,0,10">
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

                        <Label Grid.Row="2" Grid.Column="0" Content="OU (подразделение):" VerticalAlignment="Center" ToolTip="Оставьте пустым для контейнера по умолчанию"/>
                        <TextBox x:Name="txtOU" Grid.Row="2" Grid.Column="1" Margin="10,2,0,2" VerticalAlignment="Center"/>
                    </Grid>
                </GroupBox>

                <GroupBox Header="Настройка прокси (опционально)" Padding="10" Margin="0,0,0,10">
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

                <!-- Прогресс-бар (только на шаге 2) -->
                <ProgressBar x:Name="progressBar" Height="20" Margin="0,5,0,10" IsIndeterminate="False" Minimum="0" Maximum="100" Value="0"/>
            </StackPanel>
        </Grid>

        <!-- Кнопки навигации -->
        <StackPanel Grid.Row="2" HorizontalAlignment="Center" Margin="0,10,0,0" Orientation="Horizontal">
            <Button x:Name="btnBack" Content="Назад" Width="80" Height="30" Margin="0,0,10,0" IsEnabled="False"/>
            <Button x:Name="btnNext" Content="Далее" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Отмена" Width="80" Height="30" IsCancel="True"/>
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
$txtMask = $window.FindName("txtMask")             # новое поле для маски
$txtGateway = $window.FindName("txtGateway")
$txtDNS1 = $window.FindName("txtDNS1")
$txtDNS2 = $window.FindName("txtDNS2")
$txtComputerName = $window.FindName("txtComputerName")
$txtDomain = $window.FindName("txtDomain")
$txtOU = $window.FindName("txtOU")
$chkProxy = $window.FindName("chkProxy")
$txtProxyAddress = $window.FindName("txtProxyAddress")
$txtProxyPort = $window.FindName("txtProxyPort")
$progressBar = $window.FindName("progressBar")
$btnBack = $window.FindName("btnBack")
$btnNext = $window.FindName("btnNext")
$btnCancel = $window.FindName("btnCancel")
$step1 = $window.FindName("step1")
$step2 = $window.FindName("step2")
$lblStep = $window.FindName("lblStep")

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

# Установка начального состояния мастера
$lblStep.Text = "Шаг 1 из 2: Сетевые параметры"
$btnBack.IsEnabled = $false
$btnNext.Content = "Далее"
$step1.Visibility = [System.Windows.Visibility]::Visible
$step2.Visibility = [System.Windows.Visibility]::Collapsed

# -----------------------------------------------------------------------
# Функция для отображения окна подтверждения
# -----------------------------------------------------------------------
function Show-ConfirmationWindow {
    param(
        $CurrentIP, $CurrentMask, $CurrentGateway, $CurrentDNS1, $CurrentDNS2,
        $NewIP, $NewMask, $NewGateway, $NewDNS1, $NewDNS2,
        $CurrentName, $NewName, $Domain, $OU, $ProxyEnabled, $ProxyAddr, $ProxyPort
    )

    $confirmXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Подтверждение настроек" Height="450" Width="550"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="12">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Проверьте введённые параметры перед применением:" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Параметр" FontWeight="Bold" Margin="0,0,10,0"/>
                <TextBlock Grid.Row="0" Grid.Column="1" Text="Текущее значение" FontWeight="Bold" Margin="0,0,10,0"/>
                <TextBlock Grid.Row="0" Grid.Column="2" Text="Новое значение" FontWeight="Bold"/>

                <TextBlock Grid.Row="1" Grid.Column="0" Text="IP-адрес"/>
                <TextBlock x:Name="lblCurrentIP" Grid.Row="1" Grid.Column="1"/>
                <TextBlock x:Name="lblNewIP" Grid.Row="1" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="2" Grid.Column="0" Text="Маска подсети"/>
                <TextBlock x:Name="lblCurrentMask" Grid.Row="2" Grid.Column="1"/>
                <TextBlock x:Name="lblNewMask" Grid.Row="2" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="3" Grid.Column="0" Text="Шлюз"/>
                <TextBlock x:Name="lblCurrentGateway" Grid.Row="3" Grid.Column="1"/>
                <TextBlock x:Name="lblNewGateway" Grid.Row="3" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="4" Grid.Column="0" Text="Предпочитаемый DNS"/>
                <TextBlock x:Name="lblCurrentDNS1" Grid.Row="4" Grid.Column="1"/>
                <TextBlock x:Name="lblNewDNS1" Grid.Row="4" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="5" Grid.Column="0" Text="Альтернативный DNS"/>
                <TextBlock x:Name="lblCurrentDNS2" Grid.Row="5" Grid.Column="1"/>
                <TextBlock x:Name="lblNewDNS2" Grid.Row="5" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="6" Grid.Column="0" Text="Имя компьютера"/>
                <TextBlock x:Name="lblCurrentName" Grid.Row="6" Grid.Column="1"/>
                <TextBlock x:Name="lblNewName" Grid.Row="6" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="7" Grid.Column="0" Text="Домен"/>
                <TextBlock x:Name="lblDomain" Grid.Row="7" Grid.Column="1" Text="(не применимо)"/>
                <TextBlock x:Name="lblNewDomain" Grid.Row="7" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="8" Grid.Column="0" Text="OU"/>
                <TextBlock x:Name="lblOU" Grid.Row="8" Grid.Column="1" Text="(не применимо)"/>
                <TextBlock x:Name="lblNewOU" Grid.Row="8" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>

                <TextBlock Grid.Row="9" Grid.Column="0" Text="Прокси"/>
                <TextBlock x:Name="lblProxy" Grid.Row="9" Grid.Column="1" Text="(не применимо)"/>
                <TextBlock x:Name="lblNewProxy" Grid.Row="9" Grid.Column="2" FontWeight="Bold" Foreground="Green"/>
            </Grid>
        </ScrollViewer>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,15,0,0">
            <Button x:Name="btnConfirmApply" Content="Применить" Width="100" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="btnConfirmCancel" Content="Отмена" Width="100" Height="30" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $confirmReader = New-Object System.Xml.XmlNodeReader $confirmXaml
    $confirmWindow = [Windows.Markup.XamlReader]::Load($confirmReader)

    # Получаем элементы
    $lblCurrentIP = $confirmWindow.FindName("lblCurrentIP")
    $lblNewIP = $confirmWindow.FindName("lblNewIP")
    $lblCurrentMask = $confirmWindow.FindName("lblCurrentMask")
    $lblNewMask = $confirmWindow.FindName("lblNewMask")
    $lblCurrentGateway = $confirmWindow.FindName("lblCurrentGateway")
    $lblNewGateway = $confirmWindow.FindName("lblNewGateway")
    $lblCurrentDNS1 = $confirmWindow.FindName("lblCurrentDNS1")
    $lblNewDNS1 = $confirmWindow.FindName("lblNewDNS1")
    $lblCurrentDNS2 = $confirmWindow.FindName("lblCurrentDNS2")
    $lblNewDNS2 = $confirmWindow.FindName("lblNewDNS2")
    $lblCurrentName = $confirmWindow.FindName("lblCurrentName")
    $lblNewName = $confirmWindow.FindName("lblNewName")
    $lblNewDomain = $confirmWindow.FindName("lblNewDomain")
    $lblNewOU = $confirmWindow.FindName("lblNewOU")
    $lblNewProxy = $confirmWindow.FindName("lblNewProxy")
    $btnConfirmApply = $confirmWindow.FindName("btnConfirmApply")
    $btnConfirmCancel = $confirmWindow.FindName("btnConfirmCancel")

    # Заполняем значения
    $lblCurrentIP.Text = if ($CurrentIP) { $CurrentIP } else { "(не задан)" }
    $lblNewIP.Text = $NewIP
    $lblCurrentMask.Text = if ($CurrentMask) { $CurrentMask } else { "(не задан)" }
    $lblNewMask.Text = $NewMask
    $lblCurrentGateway.Text = if ($CurrentGateway) { $CurrentGateway } else { "(не задан)" }
    $lblNewGateway.Text = $NewGateway
    $lblCurrentDNS1.Text = if ($CurrentDNS1) { $CurrentDNS1 } else { "(не задан)" }
    $lblNewDNS1.Text = $NewDNS1
    $lblCurrentDNS2.Text = if ($CurrentDNS2) { $CurrentDNS2 } else { "(не задан)" }
    $lblNewDNS2.Text = if ($NewDNS2) { $NewDNS2 } else { "(не указан)" }
    $lblCurrentName.Text = $CurrentName
    $lblNewName.Text = $NewName
    $lblNewDomain.Text = $Domain
    $lblNewOU.Text = if ($OU) { $OU } else { "(по умолчанию)" }

    if ($ProxyEnabled) {
        $lblNewProxy.Text = "$ProxyAddr`:$ProxyPort"
    } else {
        $lblNewProxy.Text = "Не настроен"
    }

    # Переменная для результата
    $script:confirmResult = $false

    $btnConfirmApply.Add_Click({
        $script:confirmResult = $true
        $confirmWindow.DialogResult = $true
        $confirmWindow.Close()
    })

    $btnConfirmCancel.Add_Click({
        $script:confirmResult = $false
        $confirmWindow.DialogResult = $false
        $confirmWindow.Close()
    })

    $confirmWindow.ShowDialog() | Out-Null
    return $script:confirmResult
}

# Обработчик кнопки "Далее"
$btnNext.Add_Click({
    if ($step1.Visibility -eq [System.Windows.Visibility]::Visible) {
        # --- ШАГ 1: Валидация IP, маски, шлюза, DNS ---
        $ip = $txtIP.Text.Trim()
        $mask = $txtMask.Text.Trim()
        $gw = $txtGateway.Text.Trim()
        $dns1 = $txtDNS1.Text.Trim()
        $dns2 = $txtDNS2.Text.Trim()

        if (-not (Test-ValidIP $ip)) {
            [System.Windows.MessageBox]::Show("Некорректный IP-адрес.", "Ошибка", "OK", "Error")
            return
        }
        if (-not (Test-ValidSubnetMask $mask)) {
            [System.Windows.MessageBox]::Show("Некорректная маска подсети (пример: 255.255.255.0).", "Ошибка", "OK", "Error")
            return
        }
        if (-not (Test-ValidIP $gw)) {
            [System.Windows.MessageBox]::Show("Некорректный шлюз.", "Ошибка", "OK", "Error")
            return
        }
        if (-not (Test-ValidIP $dns1)) {
            [System.Windows.MessageBox]::Show("Некорректный адрес предпочитаемого DNS-сервера.", "Ошибка", "OK", "Error")
            return
        }
        if ($dns2 -and -not (Test-ValidIP $dns2)) {
            [System.Windows.MessageBox]::Show("Некорректный адрес альтернативного DNS-сервера (если не используется, оставьте пустым).", "Ошибка", "OK", "Error")
            return
        }

        # Переход на шаг 2
        $step1.Visibility = [System.Windows.Visibility]::Collapsed
        $step2.Visibility = [System.Windows.Visibility]::Visible
        $lblStep.Text = "Шаг 2 из 2: Компьютер и домен"
        $btnBack.IsEnabled = $true
        $btnNext.Content = "Готово"
    } else {
        # --- ШАГ 2: Валидация имени ПК, домена, прокси и проверка DC ---
        $compName = $txtComputerName.Text.Trim()
        $domain = $txtDomain.Text.Trim()
        $ou = $txtOU.Text.Trim()

        if (-not (Test-ValidComputerName $compName)) {
            [System.Windows.MessageBox]::Show("Некорректное имя компьютера (до 15 символов, только буквы, цифры, дефис).", "Ошибка", "OK", "Error")
            return
        }
        if (-not $domain) {
            [System.Windows.MessageBox]::Show("Поле 'Домен' обязательно.", "Ошибка", "OK", "Error")
            return
        }

        # Проверка прокси (если включена)
        $proxyEnabled = $chkProxy.IsChecked
        $proxyAddr = $txtProxyAddress.Text.Trim()
        $proxyPort = $txtProxyPort.Text.Trim()
        if ($proxyEnabled) {
            if (-not $proxyAddr -or -not $proxyPort) {
                [System.Windows.MessageBox]::Show("Для настройки прокси заполните адрес и порт.", "Ошибка", "OK", "Error")
                return
            }
            if (-not ($proxyPort -match '^\d+$' -and [int]$proxyPort -gt 0 -and [int]$proxyPort -le 65535)) {
                [System.Windows.MessageBox]::Show("Некорректный порт прокси.", "Ошибка", "OK", "Error")
                return
            }
        }

        # Проверка доступности контроллера домена
        if (-not (Test-DomainController -Domain $domain)) {
            $res = [System.Windows.MessageBox]::Show("Контроллер домена $domain не доступен. Продолжить?", "Предупреждение", "YesNo", "Warning")
            if ($res -eq 'No') { return }
        }

        # Получение текущих значений для окна подтверждения
        $selectedAdapterString = $cmbAdapters.SelectedItem
        $adapterIndex = [int]($selectedAdapterString -split '\(Index ')[1] -replace '\)',''
        $currentIPObj = Get-NetIPAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -ne 'WellKnown' } | Select-Object -First 1
        $currentIP = if ($currentIPObj) { $currentIPObj.IPAddress } else { $null }
        $currentPrefix = if ($currentIPObj) { $currentIPObj.PrefixLength } else { $null }
        # Преобразуем текущий префикс в маску для отображения
        $currentMask = if ($currentPrefix) {
            $bits = '1' * $currentPrefix + '0' * (32 - $currentPrefix)
            $octets = @()
            for ($i=0; $i -lt 4; $i++) {
                $octet = [Convert]::ToByte($bits.Substring($i*8, 8), 2)
                $octets += $octet
            }
            $octets -join '.'
        } else { $null }

        $currentGatewayObj = Get-NetRoute -InterfaceIndex $adapterIndex -DestinationPrefix '0.0.0.0/0' | Select-Object -First 1
        $currentGateway = if ($currentGatewayObj) { $currentGatewayObj.NextHop } else { $null }
        $currentDNS = Get-DnsClientServerAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4 | Select-Object -First 1
        $currentDNS1 = if ($currentDNS.ServerAddresses.Count -gt 0) { $currentDNS.ServerAddresses[0] } else { $null }
        $currentDNS2 = if ($currentDNS.ServerAddresses.Count -gt 1) { $currentDNS.ServerAddresses[1] } else { $null }
        $currentName = $env:COMPUTERNAME

        # Получаем новую маску для отображения (она уже введена в txtMask)
        $newMask = $txtMask.Text.Trim()

        # Показываем окно подтверждения
        $confirmResult = Show-ConfirmationWindow -CurrentIP $currentIP -CurrentMask $currentMask -CurrentGateway $currentGateway -CurrentDNS1 $currentDNS1 -CurrentDNS2 $currentDNS2 -NewIP $txtIP.Text.Trim() -NewMask $newMask -NewGateway $txtGateway.Text.Trim() -NewDNS1 $txtDNS1.Text.Trim() -NewDNS2 $txtDNS2.Text.Trim() -CurrentName $currentName -NewName $compName -Domain $domain -OU $ou -ProxyEnabled $proxyEnabled -ProxyAddr $proxyAddr -ProxyPort $proxyPort

        if ($confirmResult) {
            $window.DialogResult = $true
            $window.Close()
        } else {
            return
        }
    }
})

# Обработчик кнопки "Назад"
$btnBack.Add_Click({
    if ($step2.Visibility -eq [System.Windows.Visibility]::Visible) {
        $step2.Visibility = [System.Windows.Visibility]::Collapsed
        $step1.Visibility = [System.Windows.Visibility]::Visible
        $lblStep.Text = "Шаг 1 из 2: Сетевые параметры"
        $btnBack.IsEnabled = $false
        $btnNext.Content = "Далее"
    }
})

# Обработчик кнопки "Отмена"
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
$newMask = $txtMask.Text.Trim()
$prefixLength = Convert-MaskToPrefixLength -Mask $newMask   # преобразуем маску в длину префикса
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
Write-Host "IP: $newIP/$newMask (префикс $prefixLength), GW: $gateway, DNS: $dns1, $dns2"
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

# --- Настройка IP (используем PrefixLength) ---
Write-Host "`nНазначение статического IP..." -ForegroundColor Yellow
$progressBar.Value = 10
Get-NetIPAddress -InterfaceIndex $adapterIndex -AddressFamily IPv4 |
    Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
    Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
New-NetIPAddress -InterfaceIndex $adapterIndex -IPAddress $newIP -PrefixLength $prefixLength -DefaultGateway $gateway | Out-Null
Set-DnsClientServerAddress -InterfaceIndex $adapterIndex -ServerAddresses ($dns1, $dns2) | Out-Null
Write-Host "IP $newIP/$newMask (префикс $prefixLength), шлюз $gateway, DNS $dns1 $dns2 применены." -ForegroundColor Green
$progressBar.Value = 30

# --- Настройка прокси ---
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
Новый IP: $newIP
Новая маска: $newMask (префикс $prefixLength)
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