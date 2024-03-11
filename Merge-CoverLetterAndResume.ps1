# Install-Module PSWritePDF -Force  
# Update-Module PSWritePDF   

param(
    [Parameter(Mandatory=$True)]
    [string]$CoverLetterPath,
    [Parameter(Mandatory=$True)]
	[string]$ResumePath
)

# TODO: Check files are PDF

$coverLetterFileName = Split-Path -Path $CoverLetterPath -Leaf

# Add R char to CLR at beginning of filename
$FileNameArray = $coverLetterFileName.Split("_")
$FileNameArray[0] = "CLR"
$OutputFileName = $FileNameArray -Join "_"

# Output merged files to today's date folder
$date = Get-Date -Format "yyyy-MM-dd"
$outputPath = "./$($date)/"
$folderExists = Test-Path "./$($dateFolderName)"
if ($folderExists -ne $True) {
    New-Item -ItemType Directory -Path "./$dateFolderName"
}

Merge-PDF -InputFile $CoverLetterPath, $ResumePath -OutputFile "$($outputPath)$($OutputFileName)"