<#
.SYNOPSIS
  Отправляет в техподдержку RU-VDS запрос закрывающих документов за прошлый месяц.
  Логирует событие и результат.

.PARAMETER From
  Адрес отправителя.

.PARAMETER To
  Адрес получателя. По умолчанию support@ruvds.com.

.PARAMETER User
  Логин для SMTP (по умолчанию совпадает с From).

.PARAMETER AppPassword
  Пароль приложения почтового сервиса.

.PARAMETER LogPath
  Путь к файлу лога.

.PARAMETER SmtpServer
  Адрес SMTP-сервера. По умолчанию smtp.yandex.ru.

.PARAMETER SmtpPort
  Порт SMTP-сервера. По умолчанию 587.
#>

param(
  [Parameter(Mandatory=$true)][string]$From,
  [string]$To = 'support@ruvds.com',
  [string]$User,
  [Parameter(Mandatory=$true)][string]$AppPassword,
  [string]$LogPath = "$PSScriptRoot\ruvds-docs.log",
  [string]$SmtpServer = 'smtp.yandex.ru',
  [int]$SmtpPort = 587
)

if (-not $User) { $User = $From }

Add-Type -AssemblyName System.Globalization

function Write-Log {
  param([string]$Message)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  "$ts $Message" | Out-File -FilePath $LogPath -Encoding UTF8 -Append
}

try {
  $prev = (Get-Date).AddMonths(-1)
  $ru = [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')
  $monthTitle = $prev.ToString('MMMM yyyy', $ru)

  $Subject = "Запрос закрывающих документов за $monthTitle"

  $Body = @"
Здравствуйте.

Просим предоставить закрывающие документы за $monthTitle по нашему договору/лицевому счёту.
Готовы подтвердить данные и предоставить при необходимости доп. информацию.

Спасибо.
"@

  Write-Log "Начало отправки. Тема: '$Subject' Получатель: $To"
  $msg = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
  $msg.BodyEncoding    = [System.Text.Encoding]::UTF8
  $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
  $msg.IsBodyHtml = $false

  # копия себе
  $msg.Bcc.Add($From)

  $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
  $smtp.EnableSsl = $true
  $smtp.Credentials = New-Object System.Net.NetworkCredential($User, $AppPassword)

  $smtp.Send($msg)
  Write-Log "Успешно отправлено."
}
catch {
  Write-Log ("Ошибка отправки: " + $_.Exception.Message)
  throw
}
