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

# Read the input parameters from the text file and store them in variables
$Application_Title = Get-Content $ParameterPath | Select-Object -Index 0
$Application_Link = Get-Content $ParameterPath | Select-Object -Index 1
$Position_Name = Get-Content $ParameterPath | Select-Object -Index 2
$Company_Name = Get-Content $ParameterPath | Select-Object -Index 3
$Contact_Name = Get-Content $ParameterPath | Select-Object -Index 4
$Contact_Title = Get-Content $ParameterPath | Select-Object -Index 5
$Address = Get-Content $ParameterPath | Select-Object -Index 6
$City_Postal = Get-Content $ParameterPath | Select-Object -Index 7
$Them_Reason = Get-Content $ParameterPath | Select-Object -Index 9
$Bullet_Title_One = Get-Content $ParameterPath | Select-Object -Index 11
$Bullet_Reason_One = Get-Content $ParameterPath | Select-Object -Index 12
$Bullet_Title_Two = Get-Content $ParameterPath | Select-Object -Index 14
$Bullet_Reason_Two = Get-Content $ParameterPath | Select-Object -Index 15
$Bullet_Title_Three = Get-Content $ParameterPath | Select-Object -Index 17
$Bullet_Reason_Three = Get-Content $ParameterPath | Select-Object -Index 18


# Copy template to new folder
$dateFolderName = Get-Date -Format "yyyy-MM-dd"
$folderExists = Test-Path "./$($dateFolderName)"
if ($folderExists -ne $True) {
    New-Item -ItemType Directory -Path "./$dateFolderName"
}

$OutputPath = "$($dateFolderName)\CL_$($Company_Name)_$($Position_Name).docx"
Copy-Item -Path $TemplatePath -Destination $OutputPath
$newCl = Get-ChildItem $OutputPath
$newClFullPath = $newCl.FullName

# Read the content of the template document and store it in a variable
$Word = New-Object -ComObject word.application
$Document = $Word.Documents.Open($newClFullPath)

# Replace the placeholders with the input parameters
$Date = Get-Date -DisplayHint Date
$parameterDictionary = @{
    "<Date>" = $Date.DateTime;
    "<ApplicationTitle>" = $Application_Title;
    "<ApplicationLink>" = $Application_Link;
    "<Position>" = $Position_Name;
    "<CompanyName>" = $Company_Name;
    "<ContactName>" = $Contact_Name;
    "<ContactTitle>" = $Contact_Title;
    "<Address>" = $Address;
    "<CityAndPostalCode>" = $City_Postal;
    "<AboutThem>" = $Them_Reason;
    "<BulletTitleOne>" = $Bullet_Title_One;
    "<BulletReasonOne>" = $Bullet_Reason_One;
    "<BulletTitleTwo>" = $Bullet_Title_Two;
    "<BulletReasonTwo>" = $Bullet_Reason_Two;
    "<BulletTitleThree>" = $Bullet_Title_Three;
    "<BulletReasonThree>" = $Bullet_Reason_Three;
}

function wordSearch($Document, $existingValue, $replacingValue){
    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $true
    $wrap = $wdFindContinue
    $wdFindContinue = 1
    $Format = $false
    $ReplaceAll = 2

    $Document.Content.Find.Execute($existingValue, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, $replacingValue, $ReplaceAll)
}

foreach($key in $parameterDictionary.Keys) {
    $existingValue = $key
    $replacingValue = $parameterDictionary[$key]

    wordSearch $Document $existingValue $replacingValue
}

# Save, Close, and Quit https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
$Document.Close(-1)
$Word.Quit()

