$productKey = '<product_key_here>' # Don't include hyphen
$vb6IsoFileName = 'en_vb6_ent_cd1.iso' # optional. If unused, manually extract files to 

$currentDir = (Get-Location).Path

$vb6IsoPath = "$currentDir\$vb6IsoFileName"
$vb6SourceDir = "$currentDir\vb6"
$vb6SetupExe = "$vb6SourceDir\Setup\acmsetup.exe"
$vb6StfPath= "$currentDir\vb98ent-custom.stf"
#$vb6StfPath= "$vb6SourceDir\Setup\vb98ent.stf"
$vb6InstallLogPath = "$currentDir\vb6_install.log"

$sp6ExtractExePath = "$currentDir\vs6sp6B.exe"
$sp6SourceDir = "$currentDir\sp6"
$sp6StfPath= "$sp6SourceDir\sp698vbo.stf"
$sp6SetupExe = "$sp6SourceDir\acmsetup.exe"
$sp6InstallLogPath = "$currentDir\sp6_install.log"

$oldJavaDllPath = "$env:SystemRoot\SysWOW64\MSJAVA.DLL"

$extractedFromIso = $false

Write-Host "Beginning VB6 Enterprise install script."

if ($productKey -eq '<product_key_here>') {
    Write-Error "Product key not set. Edit this script by entering the product key and rerun."
    exit 1
}

Write-Host "currentDir: '$currentDir'"

if (!(Test-Path -Path $sp6ExtractExePath)) {
    Write-Error "Missing dependency. The script expects the Visual Basic 6 Service Pack 6 .exe at '$sp6SetupExe'"
    exit 1
}

# for extra safety -- ensures we don't delete a directory with other stuff in it during cleanup
if (Test-Path -Path $sp6SourceDir) {
    Write-Error "Please remove '$sp6SourceDir' before running this script." 
    exit 1
}

if (!(Test-Path -Path $vb6SetupExe -PathType Leaf)) {

    Write-Host "Extracting vb6 install files from .iso at '$vb6IsoPath'"

    if (!(Test-Path -Path $vb6IsoPath)) {
        Write-Error "Missing dependency. The script expects the Visual Basic 6 Enterprise .iso at '$vb6IsoPath'"
        exit 1
    }

    # for extra safety -- ensures we don't delete a directory with other stuff in it during cleanup
    if (Test-Path -Path $vb6SourceDir) {
        Write-Error "Please remove '$vb6SourceDir' before running this script."
        exit 1
    }

    $disk = Mount-DiskImage -ImagePath $vb6IsoPath -PassThru -ErrorAction Stop
    try {
        $diskVolume = Get-Volume -DiskImage $disk
        New-Item -ItemType Directory -Force -Path $vb6SourceDir
        $driveLetter = $diskVolume.DriveLetter
        Copy-Item "${driveLetter}:\*" -Recurse -Destination $vb6SourceDir -Confirm
        $extractedFromIso = $true;
    } catch {
        throw
    } finally {
        Dismount-DiskImage -ImagePath $vb6IsoPath
    }
}
else
{
    Write-Host "Found existing vb6 setup file. Skipping vb6 iso extraction."
}

# Check we have the proper paths
if (!(Test-Path -Path $vb6SourceDir)) {
    Write-Error "The path '$vb6SourceDir' was not found. Extract the files from the .iso image to that location and try again."
    exit 1
}

if (!(Test-Path -Path $vb6SetupExe -PathType Leaf)) {
    Write-Error "The file '$vb6SetupExe' was not found. Did you extract the files from the .iso image to '$vb6SourceDir'?"
    exit 1
}

# Trick setup into thinking java is up-to-date with fake 0-byte dll it checks for.
if (!(Test-Path -Path $oldJavaDllPath -PathType Leaf)) {
    New-Item -Path $oldJavaDllPath -ItemType File
}

# Add registry key to tell acmsetup that the initial setup UI doesn't need to be shown
New-Item "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic\SetupWizard" -Force -ErrorAction Stop `
    | New-ItemProperty -Name aspo -PropertyType DWord -Value 0 -ErrorAction Stop


Write-Host "Installing Visual Basic 6.0 Enterprise"
# Run the setup
# /K - Product Key
# /T - STF file that seems to be and early predecessor to the MSI format
# /S - Source Directory? ie. where the install files are located.
# /n - User's Name
# /o - User's Organization
# /b - Install Type?? Might be 1 for typical, 2 for custom. Not sure. Typical will not work.
# /G - Install log path
# /QNT - silent install without UI
$setupArgs = "/K ""$productKey"" /T ""$vb6StfPath"" /S ""$vb6SourceDir"" /G ""$vb6InstallLogPath"" /n ""MyName"" /o ""MyCompany"" /b 2  /qnt"
$process = Start-Process -PassThru -FilePath $vb6SetupExe -ArgumentList $setupArgs
$process.WaitForExit()

if (($process.ExitCode -eq 7)) {
    # REQUIRED for sp6 install in docker:
    # In a windows docker container, I get exit code 7 right before the end of the setup.
    #  I've only gotten in docker so I ignore it. ACME setup doesn't add an install entry
    #  when this error occurs. After comparing logs where we get exit 0, this seems to be
    #  the only remaining step. This registry value is pulled from a successful install
    #  on a full windows machine.
    New-Item "HKLM:\SOFTWARE\WOW6432Node\Microsoft\MS Setup (ACME)\Table Files" -Force -ErrorAction Stop `
    | New-ItemProperty -Name "Visual Basic 6.0 Enterprise Edition@v6.0.0.0.0626 (1033)" -PropertyType String -Value "C:\Program Files (x86)\Microsoft Visual Studio\VB98\Setup\1033\setup.stf" -ErrorAction Stop
} elseif (($process.ExitCode -ne 0)) {
    $ecode = $process.ExitCode
    Write-Error "Install failed with exit code '$ecode'. See the log at '$vb6InstallLogPath' for details." -ErrorAction Stop
}
$ecode = $process.ExitCode
Write-Host "Visual Basic 6 Install completed with code '$ecode'."

if ($extractedFromIso) {
    Write-Host "Cleaning up vb6 install files"
    Remove-Item -Path $sp6SourceDir -Force -Recurse
}

Write-Host "Extracting the Service Pack 6 for VB6 files"

$spExtractProcess = Start-Process -PassThru -FilePath $sp6ExtractExePath -ArgumentList "/T:""$sp6SourceDir"" /Q"
$spExtractProcess.WaitForExit()
if ($spExtractProcess.ExitCode -ne 0) {
    $ecode = $process.ExitCode
    Write-Error "Failed to extract the Service Pack 6 files. Got exit code '$ecode'." -ErrorAction Stop
}

Write-Host "Running the the VB6 Service Pack 6 for VB6 Install"

$setupArgs = "/T ""$sp6StfPath"" /S ""$sp6SourceDir"" /G ""$sp6InstallLogPath"" /qnt"
$spProcess = Start-Process -PassThru -FilePath $sp6SetupExe -ArgumentList $setupArgs
$spProcess.WaitForExit()
if ($spProcess.ExitCode -ne 0) {
    $ecode = $process.ExitCode
    Write-Error "Failed to install Service Pack 6 for VB6. Got exit code '$ecode'." -ErrorAction Stop
}

Write-Host "Cleaning up files"
Remove-Item -Path $sp6SourceDir -Force -Recurse

Write-Host "done."