<#
.SYNOPSIS
  Отправляет в техподдержку RU-VDS запрос закрывающих документов за прошлый месяц.
  Логирует событие и результат.

.PARAMETER From
  Адрес отправителя на Яндексе.

.PARAMETER To
  Адрес получателя. По умолчанию support@ruvds.com.

.PARAMETER User
  Логин для SMTP (обычно тот же, что From).

.PARAMETER AppPassword
  Пароль приложения Яндекс.Почты.

.PARAMETER LogPath
  Путь к файлу лога.
#>

param(
  [Parameter(Mandatory=$true)][string]$From,
  [string]$To = 'support@ruvds.com',
  [Parameter(Mandatory=$true)][string]$User,
  [Parameter(Mandatory=$true)][string]$AppPassword,
  [string]$LogPath = "$PSScriptRoot\ruvds-docs.log"
)


$SMTPServer='smtp.yandex.ru'
$SMTPPort=587


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
  # например: "август 2025"

  $Subject = "Запрос закрывающих документов за $monthTitle"

  $Body = @"
Здравствуйте.

Просим предоставить закрывающие документы за $monthTitle по нашему договору/лицевому счёту.
Готовы подтвердить данные и предоставить при необходимости доп. информацию.

Спасибо.
"@

  Write-Log 'Начало отправки. Тема: ''$Subject'' Получатель: $To'
  $msg = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
  $msg.BodyEncoding = [System.Text.Encoding]::UTF8
  $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
  $msg.IsBodyHtml = $false
  
  $msg.Bcc.Add($From)

  $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
  $smtp.EnableSsl = $true
  $smtp.Credentials = New-Object System.Net.NetworkCredential($User, $AppPassword)

   $smtp.Send($msg)
   Write-Log "Успешно отправлено."
}
catch {
  Write-Log ("Ошибка отправки: " + $_.Exception.Message)
  throw
}
