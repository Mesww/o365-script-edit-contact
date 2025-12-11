<#
    Description: Import/Update Contacts from CSV with Logging and Module Check
#>

# ---------------------------------------------------------------------------
# 1. การตั้งค่าตัวแปร (CONFIG)
# ---------------------------------------------------------------------------
$CsvFilePath = ".\uploads\contacts.csv"  # <-- แก้ไข path ไฟล์ CSV ของคุณที่นี่
$LogFolder   = ".\logs"          # <-- โฟลเดอร์ที่จะเก็บ Log
$LogFile     = "$LogFolder\ContactLog_$(Get-Date -Format 'yyyyMMdd-HHmm').txt"

# สร้าง Folder Log หากยังไม่มี
if (!(Test-Path -Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder | Out-Null
}

# ฟังก์ชันสำหรับเขียน Log
function Write-Log {
    Param ([string]$Message, [string]$Type = "INFO")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogContent = "[$TimeStamp] [$Type] $Message"
    Add-Content -Path $LogFile -Value $LogContent
    Write-Host $LogContent -ForegroundColor $(IF ($Type -eq "ERROR") {"Red"} elseif ($Type -eq "WARNING") {"Yellow"} else { "Green"})
}

# ---------------------------------------------------------------------------
# 2. ตรวจสอบและติดตั้ง Module (MODULE CHECK)
# ---------------------------------------------------------------------------
Write-Log "Starting Script..."
try {
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Module 'ExchangeOnlineManagement' not found. Installing..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Log "Module installed successfully."
    } else {
        Write-Log "Module 'ExchangeOnlineManagement' is already installed."
    }
}
catch {
    Write-Log "CRITICAL ERROR: Failed to install module. Details: $_" "ERROR"
    Exit
}

# ---------------------------------------------------------------------------
# 3. เชื่อมต่อ Exchange Online (CONNECT)
# ---------------------------------------------------------------------------
try {
    # เช็คว่าต่ออยู่แล้วหรือยัง ถ้ายังให้ connect
    if (!([bool](Get-ConnectionInformation -ErrorAction SilentlyContinue))) {
        Write-Log "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        Write-Log "Connected successfully."
    }
}
catch {
    Write-Log "CRITICAL ERROR: Failed to connect to Exchange Online. Details: $_" "ERROR"
    Exit
}

# ---------------------------------------------------------------------------
# 4. เริ่มกระบวนการ Import/Update (PROCESS)
# ---------------------------------------------------------------------------

# ตรวจสอบว่าไฟล์ CSV มีอยู่จริงไหม
if (!(Test-Path $CsvFilePath)) {
    # ในกรณีต้องการเลือกไฟล์ผ่าน Dialog
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = "C:\"
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.Multiselect = $false
    $dialog.Title = "Select the CSV file for Addresses"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $CsvFilePath = $dialog.FileName
    } else {
        Write-Log "No CSV file selected. Exiting." "ERROR"
        Exit
    }
}

# อ่านไฟล์ CSV
try {
    $Contacts = Import-Csv $CsvFilePath
    Write-Log "Loaded $( $Contacts.Count ) contacts from CSV."
}
catch {
    Write-Log "Error reading CSV file. Details: $_" "ERROR"
    Exit
}

$Skip_creation = $false
# ขอการยืนยันจากผู้ใช้ก่อนสร้าง Contact ใหม่
$Is_skip_confirmation = Read-Host "Do you want to skip confirm creating new contacts? (Y/N) Default(N) "
$Skip_creation = $Is_skip_confirmation -ne "Y" -and $Is_skip_confirmation -ne "y"
if (!$Skip_creation) {
    Write-Log "User chose to skip creating new contacts."
}

# วนลูปจัดการทีละคน
foreach ($Row in $Contacts) {
    $Email = $Row.Username
    $Name  = $Row."Display name"
    
    try {
        # ตรวจสอบว่า Contact นี้มีอยู่แล้วหรือไม่ (เช็คจาก Email)
        $ExistingContact = Get-Mailbox  -Identity $Email -ErrorAction SilentlyContinue

        if ($ExistingContact) {
            # --- กรณีมีอยู่แล้ว ให้ UPDATE ---
            Write-Log "Contact found: $Email. Updating..."
            
            # 1. Update ข้อมูล Mail (Set-User) - Skip if DisplayName is empty
            # if (![string]::IsNullOrWhiteSpace($Name)) {
            #     Set-User -Identity $Email -DisplayName $Name -Confirm:$Skip_creation -ErrorAction Stop
            # }
            
            # 2. Update ข้อมูลทั่วไป (Set-Contact) - Skip empty values
            $updateParams = @{
                Identity = $Email
                Confirm = $Skip_creation
                ErrorAction = "Stop"
            }
            
            if (![string]::IsNullOrWhiteSpace($Row."First name")) { $updateParams.FirstName = $Row."First name" } else { Write-Log "First name is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row."Last name")) { $updateParams.LastName = $Row."Last name" } else { Write-Log "Last name is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row.Company)) { $updateParams.Company = $Row.Company } else { Write-Log "Company is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row.Office)) { $updateParams.Office = $Row.Office } else { Write-Log "Office is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row.Department)) { $updateParams.Department = $Row.Department } else { Write-Log "Department is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row."Job title")) { $updateParams.Title = $Row."Job title" } else { Write-Log "Job title is empty, skipping update for this field." "WARNING" }
            if (![string]::IsNullOrWhiteSpace($Row.Notes)) { $updateParams.Notes = $Row.Notes } else { Write-Log "Notes is empty, skipping update for this field." "WARNING" }
            

            Set-User @updateParams

            Write-Log "SUCCESS: Updated contact '$Name' ($Email)"
        }
        else {
            # --- กรณีไม่มี ให้ CREATE NEW ---
            throw "$Email does not exist."
        }
    }
    catch {
        # บันทึก Error รายคน ลงใน Log
        Write-Log "FAILED processing '$Name' ($Email). Details: $_" "ERROR"
    }
}

Write-Log "Contact update process completed."
Write-Log "Check the log file at: $LogFile"


# ---------------------------------------------------------------------------
# 6. ตัดการเชื่อมต่อ Exchange Online (DISCONNECT)
# ---------------------------------------------------------------------------
$Is_Disconnect = Read-Host "Do you want to disconnect from Exchange Online? (Y/N) Default(N) "
$Disconnect = $Is_Disconnect -eq "Y" -or $Is_Disconnect -eq "y"

try {
    if ($Disconnect) {
        Write-Log "Disconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Disconnected successfully."
    }
    else {
        Write-Log "User chose to remain connected to Exchange Online."
    }
}
catch {
    Write-Log "Error disconnecting from Exchange Online. Details: $_" "ERROR"
}