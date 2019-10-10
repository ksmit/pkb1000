#
# pkb1000/install-shortcut
# Installs the Journal shortcut. After being set up, Windows should let you invoke
# it with the system-wide hotkey CTRL + ALT + J
#

$journalScript = $PSScriptRoot + "\journal.ps1";
$shortcutLocation = $env:APPDATA + "\Microsoft\Windows\Start Menu\Programs\pkb1000-journal.lnk";
$targetPath = "$($env:SystemRoot)\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File " + '"' + $journalScript + '"'

if ((test-path $shortcutLocation)) {
    Remove-Item $shortcutLocation
}

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutLocation)
$Shortcut.TargetPath = '' + $env:SystemRoot + '\system32\WindowsPowerShell\v1.0\powershell.exe'
$Shortcut.Arguments = '-ExecutionPolicy Bypass -File ' + '"' + $journalScript + '"'
$Shortcut.Hotkey = 'CTRL+ALT+J'
$Shortcut.Save()
