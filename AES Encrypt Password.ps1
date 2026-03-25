<#
.SYNOPSIS
    Creates an AES encryption key and encrypts a password for use in automated scripts.

.DESCRIPTION
    Generates a 256-bit AES key file and uses it to encrypt a password into a separate file.
    Both files are required at runtime to decrypt the credential:

        $password = Get-Content "Password.txt" | ConvertTo-SecureString -Key (Get-Content "AesKey.txt")
        $credential = New-Object System.Management.Automation.PSCredential ("username", $password)

    Run this script ONCE during setup, then store the output files in your secure directory.

    IMPORTANT: The AES key file is the secret. Anyone with access to both files can
    decrypt the password. Restrict NTFS permissions on the secure directory to only
    the service account that runs your reports.

.EXAMPLE
    .\New-SecureCredential.ps1

.NOTES
    Author:  Your Name
    Version: 1.0

    Output files (default paths - update $outputDir below):
    - C:\Scripts\Secure\AesKey.txt        — 256-bit AES key
    - C:\Scripts\Secure\Password.txt      — AES-encrypted password
#>

# ============================================
# Configuration
# ============================================
$outputDir = "C:\Scripts\Secure"
$keyFileName = "AesKey.txt"
$passFileName = "Password.txt"

$keyFilePath = Join-Path $outputDir $keyFileName
$passFilePath = Join-Path $outputDir $passFileName

# ============================================
# Create output directory if it doesn't exist
# ============================================
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "✓ Created directory: $outputDir" -ForegroundColor Green
}

# ============================================
# Safety check — don't overwrite existing files
# ============================================
if ((Test-Path $keyFilePath) -or (Test-Path $passFilePath)) {
    Write-Host "⚠ One or both output files already exist:" -ForegroundColor Yellow
    if (Test-Path $keyFilePath) { Write-Host "    $keyFilePath" -ForegroundColor Yellow }
    if (Test-Path $passFilePath) { Write-Host "    $passFilePath" -ForegroundColor Yellow }
    Write-Host ""

    $confirm = Read-Host "Overwrite? (y/N)"
    if ($confirm -ne 'y') {
        Write-Host "Aborted. No files were changed." -ForegroundColor Cyan
        exit 0
    }
}

# ============================================
# Step 1: Generate AES Key
# ============================================
Write-Host "`nStep 1: Generating 256-bit AES key..." -ForegroundColor Cyan

$aesKey = New-Object byte[] 32
[System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($aesKey)

$aesKey | Out-File -FilePath $keyFilePath -Force
Write-Host "✓ AES key saved to: $keyFilePath" -ForegroundColor Green

# ============================================
# Step 2: Prompt for password and encrypt
# ============================================
Write-Host "`nStep 2: Enter the password to encrypt" -ForegroundColor Cyan

$securePassword = Read-Host -AsSecureString -Prompt "Password"

# Verify it
$confirmPassword = Read-Host -AsSecureString -Prompt "Confirm password"

# Compare the two
$bstr1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
$bstr2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmPassword)
$plain1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr1)
$plain2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr2)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2)

if ($plain1 -ne $plain2) {
    # Clear plaintext from memory
    $plain1 = $null
    $plain2 = $null
    Write-Host "✗ Passwords do not match. No files were written." -ForegroundColor Red
    # Clean up the key file we already wrote
    if (Test-Path $keyFilePath) { Remove-Item $keyFilePath -Force }
    exit 1
}

$plain1 = $null
$plain2 = $null

# Encrypt with AES key and save
$encryptedPassword = $securePassword | ConvertFrom-SecureString -Key $aesKey
$encryptedPassword | Out-File -FilePath $passFilePath -Force
Write-Host "✓ Encrypted password saved to: $passFilePath" -ForegroundColor Green

# ============================================
# Step 3: Verify decryption works
# ============================================
Write-Host "`nStep 3: Verifying decryption..." -ForegroundColor Cyan

try {
    $testKey = Get-Content $keyFilePath
    $testPassword = Get-Content $passFilePath | ConvertTo-SecureString -Key $testKey
    $testCred = New-Object System.Management.Automation.PSCredential ("test", $testPassword)

    # Verify we can extract the plaintext (proves the round-trip works)
    $null = $testCred.GetNetworkCredential().Password
    Write-Host "✓ Decryption verified successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Decryption verification failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================
# Summary
# ============================================
Write-Host "`n========================================" -ForegroundColor Green
Write-Host "✓ Credential files created successfully" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Files:" -ForegroundColor Cyan
Write-Host "  Key:      $keyFilePath" -ForegroundColor White
Write-Host "  Password: $passFilePath" -ForegroundColor White
Write-Host ""
Write-Host "Usage in your scripts:" -ForegroundColor Cyan
Write-Host '  $password   = Get-Content "' -NoNewline -ForegroundColor White
Write-Host $passFilePath -NoNewline -ForegroundColor Yellow
Write-Host '" | ConvertTo-SecureString -Key (Get-Content "' -NoNewline -ForegroundColor White
Write-Host $keyFilePath -NoNewline -ForegroundColor Yellow
Write-Host '")' -ForegroundColor White
Write-Host '  $credential = New-Object System.Management.Automation.PSCredential ("your_username", $password)' -ForegroundColor White
Write-Host ""
Write-Host "⚠ SECURITY:" -ForegroundColor Yellow
Write-Host "  - Restrict NTFS permissions on $outputDir to the service account only" -ForegroundColor Yellow
Write-Host "  - Do NOT commit these files to source control" -ForegroundColor Yellow
Write-Host "  - Add the secure directory to your .gitignore" -ForegroundColor Yellow