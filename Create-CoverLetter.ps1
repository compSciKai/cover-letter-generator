param(
    [Parameter(Mandatory=$True)]
    [string]$TemplatePath,
    [Parameter(Mandatory=$True)]
	[string]$ParameterPath
)

$parameterPathExists = Test-Path -Path $ParameterPath
if($parameterPathExists -ne $True) {
    Write-Host "Error: Parameters not found at parameter path."
    exit
}

function Create-ClParameterDictionary($ParameterPath) {
    # Read the input parameters from the parameter file and store them in dictionary

    $Date = Get-Date -DisplayHint Date
     return @{
        "<Date>" = $Date.DateTime;
        "<ApplicationTitle>" = Get-Content $ParameterPath | Select-Object -Index 0;
        "<ApplicationLink>" = Get-Content $ParameterPath | Select-Object -Index 1;
        "<Position>" = Get-Content $ParameterPath | Select-Object -Index 2;
        "<CompanyName>" = Get-Content $ParameterPath | Select-Object -Index 3;
        "<ContactName>" = Get-Content $ParameterPath | Select-Object -Index 4;
        "<ContactTitle>" = Get-Content $ParameterPath | Select-Object -Index 5;
        "<Address>" = Get-Content $ParameterPath | Select-Object -Index 6;
        "<CityAndPostalCode>" = Get-Content $ParameterPath | Select-Object -Index 7;
        "<AboutThem>" = Get-Content $ParameterPath | Select-Object -Index 9;
        "<BulletTitleOne>" = Get-Content $ParameterPath | Select-Object -Index 11;
        "<BulletReasonOne>" = Get-Content $ParameterPath | Select-Object -Index 12;
        "<BulletTitleTwo>" = Get-Content $ParameterPath | Select-Object -Index 14;
        "<BulletReasonTwo>" = Get-Content $ParameterPath | Select-Object -Index 15;
        "<BulletTitleThree>" = Get-Content $ParameterPath | Select-Object -Index 17;
        "<BulletReasonThree>" = Get-Content $ParameterPath | Select-Object -Index 18;
        "<BulletTitleFour>" = Get-Content $ParameterPath | Select-Object -Index 20;
        "<BulletReasonFour>" = Get-Content $ParameterPath | Select-Object -Index 21;
    }
}

$parameterDictionary = Create-ClParameterDictionary $ParameterPath

# Copy template to new folder
$dateFolderName = Get-Date -Format "yyyy-MM-dd"
$folderExists = Test-Path "./$($dateFolderName)"
if ($folderExists -ne $True) {
    New-Item -ItemType Directory -Path "./$dateFolderName"
}

$newClFullPath = $null
$OutputPath = $null
try {
    $Company_Name = $parameterDictionary."<CompanyName>"
    $Position_Name = $parameterDictionary."<Position>"
    $OutputPath = "$($dateFolderName)\CL_$($Company_Name)_$($Position_Name).docx"
    Copy-Item -Path $TemplatePath -Destination $OutputPath
    $newCl = Get-ChildItem $OutputPath
    $newClFullPath = $newCl.FullName
} catch {
    Write-Host "Error: cover letter not copied to path $($newClFullPath)"
    exit
}

# Read the content of the template document and store it in a variable
$Word = New-Object -ComObject word.application
$Document = $Word.Documents.Open($newClFullPath)




function wordSearch($Document, $existingValue, $replacingValue){
    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $true
    $wdFindContinue = 1
    $wrap = $wdFindContinue
    $Format = $false
    $ReplaceAll = 2

    $Document.Content.Find.Execute(
        $existingValue, 
        $MatchCase, 
        $MatchWholeWord, 
        $MatchWildcards, 
        $MatchSoundsLike, 
        $MatchAllWordForms, 
        $Forward, 
        $wrap, 
        $Format, 
        $replacingValue, 
        $ReplaceAll
    ) | Out-Null
}

foreach($key in $parameterDictionary.Keys) {
    $existingValue = $key
    $replacingValue = $parameterDictionary[$key]

    wordSearch $Document $existingValue $replacingValue
}

# Save, Close, and Quit https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
try {
    $Document.Close(-1)
    $Word.Quit()
    Write-Host "Successfully created new cover letter: ..\$OutputPath"
} catch {
    Write-Host "Error: Could not save cover letter. Existing documents might still be open."
}


