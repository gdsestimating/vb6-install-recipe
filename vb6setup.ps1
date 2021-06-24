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

Write-Host "Beginning VB6 Enterprise install script."

if ($productKey = '<product_key_here>') {
    Write-Error "Product key not set. Edit this script by entering the product key and rerun."
}

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
if ($process.ExitCode -ne 0) {
    Write-Error "Install failed. See the log at '$logPath' for details." -ErrorAction Stop
}
Write-Host "Visual Basic 6 Install completed."

Write-Host "Extracting the VB6 Service Pack 6 files"

$spExtractProcess = Start-Process -PassThru -FilePath $sp6ExtractExePath -ArgumentList "/T:$sp6SourceDir /Q"
$spExtractProcess.WaitForExit()
if ($spExtractProcess.ExitCode -ne 0) {
    Write-Error "Failed to extract the files." -ErrorAction Stop
}

Write-Host "Running the the VB6 Service Pack 6 Install"

$setupArgs = "/T ""$sp6StfPath"" /S ""$sp6SourceDir"" /G ""$sp6InstallLogPath"" /qnt"
$spProcess = Start-Process -PassThru -FilePath $sp6SetupExe -ArgumentList $setupArgs
$spProcess.WaitForExit()
if ($spProcess.ExitCode -ne 0) {
    Write-Error "Failed to install the service pack." -ErrorAction Stop
}

Write-Host "Cleaning up files"
Remove-Item -Path $vb6SourceDir,$sp6SourceDir -Force -Recurse

Write-Host "done."