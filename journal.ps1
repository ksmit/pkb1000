#
# pkb1000/journal
# Open today's journal file. Copies in a starter template if one exists.
#

$datePath = $PSScriptRoot + "\" + (Get-Date).toString("yyyy\\MM")
$journalFilePath = $datePath + "\journal_" + (Get-Date).toString("dd") + ".md"
$journalTemplatePath = $PSScriptRoot + "\template.md"

if (-not (test-path $datePath)) {
    new-item -path $datePath -ItemType Directory
}

if (-not (test-path $journalFilePath)) {
    if ((test-path $journalTemplatePath)) {
        Copy-Item $journalTemplatePath $journalFilePath
    }
    else {
        New-Item $journalFilePath
        Set-Content $journalFilePath "Set up a template at $($journalTemplatePath) for more interesting starter content."
    }
}

code $datePath
