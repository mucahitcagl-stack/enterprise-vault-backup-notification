#Günlük olarak otomatik mail backup sonrası güncel dosya içerisindeki bilgileri alarak mail gönderiminin yapıldığı scripttir.

Add-Type -AssemblyName System.Web

#Bu alana Enterprise Vault uygulamasının kurulu olduğu sürücü üzerinde hergün çalıştıktan sonra raporları yazdığı alan yazılmalı. Kaç tane task var ise o kadar çoğaltılabilir.
#Path üzerinde task adı yazan yere EV üzerinde oluşturulan task adını yazmanız gerekmektedir.

$ReportRootPaths = @(
    "X:\Program Files (x86)\Enterprise Vault\Reports\Exchange Mailbox Archiving\Task-Adı\Scheduled",
    "X:\Program Files (x86)\Enterprise Vault\Reports\Exchange Mailbox Archiving\Task-Adı\Scheduled"
)

# Mail Ayarları
$SmtpServer = "SMTP Server Bilgisi."
$From       = "Göndericiadresi@xxxx.xxx.xx"
$To         = "aliciadresi1@xxx.xxx.xx", "aliciadresi2@xxx.xxx.xx", "aliciadresi3@xxx.xxx.xx"

$Today   = Get-Date -Format "yyyyMMdd"
$Subject = "Enterprise Vault Archiving Report - $(Get-Date -Format 'dd.MM.yyyy')"

$Attachments     = @()
$Warnings        = @()
$ReportSummaries = @()

function Get-SafeValue {
    param(
        [array]$Values,
        [int]$Index
    )

    if ($Values.Count -gt $Index -and -not [string]::IsNullOrWhiteSpace($Values[$Index])) {
        return $Values[$Index]
    }

    return "N/A"
}

function Get-EVReportValuesFromHtml {
    param(
        [string]$FilePath
    )

    $Html = Get-Content -Path $FilePath -Raw -Encoding Default

    $Matches = [regex]::Matches(
        $Html,
        '<td\s+class=["'']summaryvalue["'']\s*>(.*?)</td>',
        'IgnoreCase,Singleline'
    )

    $Values = @()

    foreach ($Match in $Matches) {
        $Value = $Match.Groups[1].Value
        $Value = $Value -replace '<[^>]+>', ''
        $Value = [System.Web.HttpUtility]::HtmlDecode($Value).Trim()
        $Values += $Value
    }

    return @{
        "Run status"                     = Get-SafeValue -Values $Values -Index 0
        "Exchange Server"                = Get-SafeValue -Values $Values -Index 1
        "Enterprise Vault site"          = Get-SafeValue -Values $Values -Index 2
        "Start time"                     = Get-SafeValue -Values $Values -Index 3
        "End time"                       = Get-SafeValue -Values $Values -Index 4
        "Processing time"                = Get-SafeValue -Values $Values -Index 5
        "Run type"                       = Get-SafeValue -Values $Values -Index 6
        "Run mode"                       = Get-SafeValue -Values $Values -Index 7
        "Run ID"                         = Get-SafeValue -Values $Values -Index 8
        "Items archived"                 = Get-SafeValue -Values $Values -Index 9

        "Mailboxes targeted"             = Get-SafeValue -Values $Values -Index 10
        "Mailboxes completed"            = Get-SafeValue -Values $Values -Index 11
        "Mailboxes partially processed"  = Get-SafeValue -Values $Values -Index 12
        "Mailboxes with no action taken" = Get-SafeValue -Values $Values -Index 13
        "Mailboxes not processed"        = Get-SafeValue -Values $Values -Index 14
        "Mailboxes with warnings"        = Get-SafeValue -Values $Values -Index 15
    }
}

foreach ($RootPath in $ReportRootPaths) {

    if (!(Test-Path $RootPath)) {
        $Warnings += "Path bulunamadı: $RootPath"
        continue
    }

    $TodayFolder = Get-ChildItem -Path $RootPath -Directory -Filter "Scheduled_$Today*" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $TodayFolder) {
        $Warnings += "BUGÜNE ait Scheduled klasörü bulunamadı: $RootPath - Lütfen backup scheduler kontrol edin."
        continue
    }

    $TodayReport = Get-ChildItem -Path $TodayFolder.FullName -File -Filter "Full-*.htm" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $TodayReport) {
        $Warnings += "BUGÜNE ait Full rapor bulunamadı: $($TodayFolder.FullName) - Lütfen backup scheduler kontrol edin."
        continue
    }

    $Attachments += $TodayReport.FullName

    $Values = Get-EVReportValuesFromHtml -FilePath $TodayReport.FullName

    $Summary = [PSCustomObject]@{
        ReportFile                  = $TodayReport.Name
        ExchangeServer              = $Values["Exchange Server"]
        EnterpriseVaultSite         = $Values["Enterprise Vault site"]
        RunStatus                   = $Values["Run status"]
        StartTime                   = $Values["Start time"]
        EndTime                     = $Values["End time"]
        ProcessingTime              = $Values["Processing time"]
        RunType                     = $Values["Run type"]
        RunMode                     = $Values["Run mode"]
        RunID                       = $Values["Run ID"]
        ItemsArchived               = $Values["Items archived"]

        MailboxesTargeted           = $Values["Mailboxes targeted"]
        MailboxesCompleted          = $Values["Mailboxes completed"]
        MailboxesPartiallyProcessed = $Values["Mailboxes partially processed"]
        MailboxesWithNoActionTaken  = $Values["Mailboxes with no action taken"]
        MailboxesNotProcessed       = $Values["Mailboxes not processed"]
        MailboxesWithWarnings       = $Values["Mailboxes with warnings"]
    }

    $ReportSummaries += $Summary
}

$Body = @"
<html>
<body style="font-family:Calibri,Arial;font-size:11pt;color:#222;">

<p>Merhaba,</p>
<p>Enterprise Vault archiving task kontrolü tamamlanmıştır.</p>

<h2 style="color:#2F5597;">📦 Enterprise Vault Rapor Özeti</h2>
"@

foreach ($Summary in $ReportSummaries) {

    $StatusIcon = "✅"
    if ($Summary.RunStatus -ne "Completed") {
        $StatusIcon = "⚠️"
    }

    $WarningIcon = "✅"
    if ($Summary.MailboxesWithWarnings -ne "0" -and $Summary.MailboxesWithWarnings -ne "N/A") {
        $WarningIcon = "⚠️"
    }

    $Body += @"
<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;margin-bottom:18px;border:1px solid #d9d9d9;">
<tr style="background-color:#2F5597;color:white;">
<td colspan="4" style="font-size:13pt;font-weight:bold;">
📄 $($Summary.ReportFile)
</td>
</tr>

<tr>
<td style="width:25%;background-color:#f3f6fa;font-weight:bold;">🖥️ Exchange Server</td>
<td style="width:25%;">$($Summary.ExchangeServer)</td>
<td style="width:25%;background-color:#f3f6fa;font-weight:bold;">🏢 EV Site</td>
<td style="width:25%;">$($Summary.EnterpriseVaultSite)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">$StatusIcon Run Status</td>
<td>$($Summary.RunStatus)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">▶️ Run Type</td>
<td>$($Summary.RunType)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">⏱️ Start Time</td>
<td>$($Summary.StartTime)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">🏁 End Time</td>
<td>$($Summary.EndTime)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">⌛ Processing Time</td>
<td>$($Summary.ProcessingTime)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">🆔 Run ID</td>
<td>$($Summary.RunID)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">📦 Items Archived</td>
<td>$($Summary.ItemsArchived)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">📬 Mailboxes Targeted</td>
<td>$($Summary.MailboxesTargeted)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">✅ Mailboxes Completed</td>
<td>$($Summary.MailboxesCompleted)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">➖ No Action Taken</td>
<td>$($Summary.MailboxesWithNoActionTaken)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">⏳ Partially Processed</td>
<td>$($Summary.MailboxesPartiallyProcessed)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">❌ Not Processed</td>
<td>$($Summary.MailboxesNotProcessed)</td>
</tr>

<tr>
<td style="background-color:#f3f6fa;font-weight:bold;">$WarningIcon Mailboxes With Warnings</td>
<td>$($Summary.MailboxesWithWarnings)</td>
<td style="background-color:#f3f6fa;font-weight:bold;">⚙️ Run Mode</td>
<td>$($Summary.RunMode)</td>
</tr>
</table>
"@
}

if ($Warnings.Count -gt 0) {
    $Body += "<h3 style='color:#C00000;'>⚠️ Uyarılar</h3><ul>"

    foreach ($Warning in $Warnings) {
        $Body += "<li>$Warning</li>"
    }

    $Body += "</ul>"
}

if ($Attachments.Count -gt 0) {
    $Body += "<p>📎 Full HTML raporları ekte iletilmiştir.</p>"
}
else {
    $Body += "<p><b>⚠️ Güncel tarihli rapor bulunamadı. Lütfen backup scheduler kontrol edin.</b></p>"
}

$Body += @"
<p>İyi çalışmalar.</p>

</body>
</html>
"@

Send-MailMessage `
    -SmtpServer $SmtpServer `
    -From $From `
    -To $To `
    -Subject $Subject `
    -Body $Body `
    -BodyAsHtml `
    -Attachments $Attachments `
    -Encoding UTF8

Write-Host "Mail gönderildi."