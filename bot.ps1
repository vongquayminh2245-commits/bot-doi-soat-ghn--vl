param(
    [string]$TargetChatId = "-1003837960650" # Default is Room 1
)

# ============================================================
# BOT TELEGRAM DOI SOAT GHN & VAN LOI
# Tu dong gui bao cao
# ============================================================

$TOKEN     = "8976128058:AAGemPMY0HruH-3VvUbEdT-FH2LTpCNa4Pc"
$BASE_URL  = "https://docs.google.com/spreadsheets/d/1uRumC5gWm31VcS3X7VzRyIYR1p-PgR46AzsMg8oPlTI/export?format=csv&gid="
$GID_S1    = "1379748291"   # Sheet 1: Van Loi & GHN Doi Soat
$GID_S10   = "741237788"    # Sheet 10: Data MCO Van Loi

# -----------------------------------------------
# Ham gui tin nhan Telegram
# -----------------------------------------------
function Send-TelegramMessage {
    param([string]$msg)
    
    # Chia tin nhan neu qua dai (gioi han 4096 ky tu cua Telegram)
    $maxLength = 3800
    $chunks = @()
    
    while ($msg.Length -gt 0) {
        if ($msg.Length -le $maxLength) {
            $chunks += $msg
            break
        }
        
        # Tim vi tri cat tai ky tu xuong dong gan nhat
        $splitIndex = $msg.LastIndexOf("`n", $maxLength)
        if ($splitIndex -lt 0) { $splitIndex = $maxLength }
        
        $chunks += $msg.Substring(0, $splitIndex)
        $msg = $msg.Substring($splitIndex).TrimStart("`n")
    }

    foreach ($chunk in $chunks) {
        $body = @{
            chat_id    = $TargetChatId
            text       = $chunk
            parse_mode = "HTML"
        }
        try {
            Invoke-RestMethod -Uri "https://api.telegram.org/bot$TOKEN/sendMessage" `
                              -Method Post `
                              -Body $body | Out-Null
            Start-Sleep -Milliseconds 500 # Tranh bi rate limit
        } catch {
            Write-Host "Loi gui Telegram: $_"
        }
    }
}

# -----------------------------------------------
# Ham lay du lieu tu Google Sheets (CSV)
# -----------------------------------------------
function Get-SheetData {
    param([string]$gid)
    $url  = $BASE_URL + $gid
    $file = "$env:TEMP\sheet_$gid.csv"
    try {
        Invoke-WebRequest -Uri $url -OutFile $file -UseBasicParsing | Out-Null
        $data = Import-Csv $file
        return $data
    } catch {
        Write-Host "Loi tai sheet GID=$gid : $_"
        return $null
    }
}

function Get-Col {
    param($row, [int]$index)
    $props = @($row.PSObject.Properties)
    if ($index -lt $props.Count) {
        $val = $props[$index].Value
        if ($val) { return $val.ToString().Trim() }
    }
    return ""
}

# ============================================================
# MAIN: Chay bao cao
# ============================================================
Write-Host "=== BOT DOI SOAT GHN & VAN LOI ==="
Write-Host "Dang tai du lieu S1 va S10..."

$s1 = Get-SheetData -gid $GID_S1
$s10 = Get-SheetData -gid $GID_S10

if ($null -eq $s1 -or $null -eq $s10) {
    Send-TelegramMessage "Loi: Khong tai duoc du lieu tu Google Sheets."
    exit 1
}

# Build map for S10 (Data MCO Van Loi) key = Ma Chuyen (Index 0)
$s10_map = @{}
foreach ($row10 in $s10) {
    $maChuyen10 = Get-Col $row10 0
    if ($maChuyen10) {
        $s10_map[$maChuyen10] = $row10
    }
}

$total_checked = 0
$issues_odo_by_kho = @{}
$total_odo_issues = 0
$issues_bks = @()
$issues_tuyen = @()
$issues_ngay = @()
$issues_kho = @()
$latest_date = [datetime]::MinValue

foreach ($row1 in $s1) {
    $maChuyen1 = Get-Col $row1 11 # MA CHUYEN DI TU HT

    if ($maChuyen1 -and $s10_map.ContainsKey($maChuyen1)) {
        $total_checked++
        $row10 = $s10_map[$maChuyen1]

        # S1 Cols
        $s1_bks   = Get-Col $row1 3
        $s1_kho   = Get-Col $row1 0
        $s1_tuyen = Get-Col $row1 1
        $s1_ngay  = Get-Col $row1 2
        
        try {
            $parsedDate = [datetime]::ParseExact($s1_ngay.Trim(), 'd/M/yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
            if ($parsedDate -gt $latest_date) {
                $latest_date = $parsedDate
            }
        } catch {}

        $s1_vao   = Get-Col $row1 8
        $s1_ra    = Get-Col $row1 9
        $s1_km    = Get-Col $row1 10

        # S10 Cols
        $s10_bks  = Get-Col $row10 1
        $s10_kho  = Get-Col $row10 4
        $s10_tuyen= Get-Col $row10 6
        $s10_ngay = Get-Col $row10 5
        $s10_bks2 = Get-Col $row10 12 # Bien so (NVXL Dien)
        $s10_di   = Get-Col $row10 9
        $s10_ve   = Get-Col $row10 10
        $s10_km   = Get-Col $row10 11

        # Kiem tra Kho
        if ($s1_kho -ne $s10_kho -and $s1_kho -and $s10_kho) {
            $issues_kho += "  - <b>$maChuyen1</b> (Ngày: $s1_ngay)`n    &#x21B3; GHN: $s10_kho | VL: $s1_kho"
        }

        # Kiem tra Tuyen
        if ($s1_tuyen -ne $s10_tuyen -and $s1_tuyen -and $s10_tuyen) {
            $issues_tuyen += "  - <b>$maChuyen1</b> (Ngày: $s1_ngay)`n    &#x21B3; GHN: $s10_tuyen | VL: $s1_tuyen"
        }

        # Kiem tra Ngay (chi bao loi neu khac nhau thuc su ve dinh dang/ngay)
        if ($s1_ngay -ne $s10_ngay -and $s1_ngay -and $s10_ngay) {
            $issues_ngay += "  - <b>$maChuyen1</b>`n    &#x21B3; GHN: $s10_ngay | VL: $s1_ngay"
        }

        # Kiem tra Bien So Xe (S10 Bien So Xe vs S1 Bien so xe OR S10 NVXL vs S1 Bien so xe)
        if ($s1_bks -ne $s10_bks -or $s1_bks -ne $s10_bks2) {
            $issues_bks += "  - <b>$maChuyen1</b> (Ngày: $s1_ngay)`n    &#x21B3; GHN: $s10_bks ($s10_bks2) | VL: $s1_bks"
        }

        # Chuan hoa so: xoa dau cham, dau phay truoc khi so sanh
        $s1_vao_num = $s1_vao -replace '[.,]',''
        $s1_ra_num = $s1_ra -replace '[.,]',''
        $s1_km_num = $s1_km -replace '[.,]',''
        
        $s10_di_num = $s10_di -replace '[.,]',''
        $s10_ve_num = $s10_ve -replace '[.,]',''
        $s10_km_num = $s10_km -replace '[.,]',''

        $lech_di = ($s1_vao_num -ne $s10_di_num)
        $lech_ve = ($s1_ra_num -ne $s10_ve_num)
        $lech_tong = ($s1_km_num -ne $s10_km_num)

        if ($lech_di -or $lech_ve -or $lech_tong) {
            $issue_str = "($s1_ngay) Tuyến $s1_tuyen (BKS: $s1_bks)`n      &#x21B3; <b>[GHN]</b> ($s10_di_num - $s10_ve_num = $s10_km_num) | <b>[VL]</b> ($s1_vao_num - $s1_ra_num = $s1_km_num)"
            
            $k = $s1_kho
            if ([string]::IsNullOrWhiteSpace($k)) { $k = "KHO KHÁC" }
            if (-not $issues_odo_by_kho.ContainsKey($k)) {
                $issues_odo_by_kho[$k] = @()
            }
            $issues_odo_by_kho[$k] += $issue_str
            $total_odo_issues++
        }
    }
}

$today = Get-Date -Format "dd/MM/yyyy"
$lines = @()
$lines += "&#x1F6A6; <b>BÁO CÁO CHÉO: DATA MCĐ vs ĐỐI SOÁT GHN</b>"
$lines += "&#x1F5D3;&#xFE0F; <b>Ngày cập nhật: $today</b>"
$lines += "&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;&#x2500;"
$lines += ""
$lines += "<b>I. TỔNG QUAN</b>"
$lines += "&#x1F537; Tổng chuyến đã đối chiếu: <b>$total_checked chuyến</b>"
if ($latest_date -ne [datetime]::MinValue) {
    $lines += "&#x1F537; Dữ liệu đối chiếu đến ngày: <b>$($latest_date.ToString('dd/MM/yyyy'))</b>"
}
$lines += ""

$lines += "<b>II. LỆCH BIỂN SỐ XE (BKS / NVXL Điền)</b>"
if ($issues_bks.Count -eq 0) { $lines += "&#x2705; Tất cả đều khớp." }
else { 
    $lines += "&#x274C; <b>$($issues_bks.Count)</b> chuyến bị lệch:"
    for ($i=0; $i -lt $issues_bks.Count; $i++) { $lines += "  $($i+1). " + $issues_bks[$i] }
}
$lines += ""

$lines += "<b>III. LỆCH SỐ KM (ODO vào/ra/tổng)</b>"
if ($total_odo_issues -eq 0) { $lines += "&#x2705; Tất cả đều khớp." }
else { 
    $lines += "&#x26A0;&#xFE0F; <b>$total_odo_issues</b> chuyến bị lệch:"
    foreach ($kho in ($issues_odo_by_kho.Keys | Sort-Object)) {
        $lines += "  &#x1F3ED; <b>KHO: $kho</b>"
        $kho_issues = $issues_odo_by_kho[$kho]
        for ($i=0; $i -lt $kho_issues.Count; $i++) { 
            $lines += "    $($i+1). " + $kho_issues[$i] 
        }
    }
}
$lines += ""

$lines += "<b>IV. LỆCH TUYẾN / KHO / NGÀY</b>"
$hasInfoLech = $false
if ($issues_kho.Count -gt 0) { $lines += "<b>* KHO ($($issues_kho.Count)):</b>"; for ($i=0; $i -lt $issues_kho.Count; $i++) { $lines += "  $($i+1). " + $issues_kho[$i] } ; $hasInfoLech=$true }
if ($issues_tuyen.Count -gt 0) { $lines += "<b>* TUYẾN ($($issues_tuyen.Count)):</b>"; for ($i=0; $i -lt $issues_tuyen.Count; $i++) { $lines += "  $($i+1). " + $issues_tuyen[$i] } ; $hasInfoLech=$true }
if ($issues_ngay.Count -gt 0) { $lines += "<b>* NGÀY ($($issues_ngay.Count)):</b>"; for ($i=0; $i -lt $issues_ngay.Count; $i++) { $lines += "  $($i+1). " + $issues_ngay[$i] } ; $hasInfoLech=$true }
if (-not $hasInfoLech) { $lines += "&#x2705; Tất cả đều khớp." }

$report = $lines -join "`n"

Send-TelegramMessage -msg $report
Write-Host "Da gui bao cao qua Telegram cho Chat ID $TargetChatId thanh cong!"


