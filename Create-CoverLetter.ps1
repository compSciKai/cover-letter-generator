# Set the path of the template document
$clTemplatePath = "./CoverLetterTemplate.docx"

# Copy template to new path
$newClRelativePath = "./CoverLetterExample.docx"
Copy-Item -Path $clTemplatePath -Destination $newClRelativePath
$newCl = Get-ChildItem $newClRelativePath
$newClFullPath = $newCl.FullName

# Set the path of the text file with the input parameters
$clParametersPath = "./InputParameters.txt"

# Read the input parameters from the text file and store them in variables
$Position_Name = Get-Content $clParametersPath | Select-Object -Index 0
# $Company_Name = Get-Content $clParametersPath | Select-Object -Index 1
# $Years_Of_Experience = Get-Content $clParametersPath | Select-Object -Index 2
# $Field_Of_Expertise = Get-Content $clParametersPath | Select-Object -Index 3
# $Skill_1 = Get-Content $clParametersPath | Select-Object -Index 4
# $Skill_2 = Get-Content $clParametersPath | Select-Object -Index 5
# $Skill_3 = Get-Content $clParametersPath | Select-Object -Index 6
# $Example_1 = Get-Content $clParametersPath | Select-Object -Index 7
# $Example_2 = Get-Content $clParametersPath | Select-Object -Index 8
# $Example_3 = Get-Content $clParametersPath | Select-Object -Index 9
# $Achievement_1 = Get-Content $clParametersPath | Select-Object -Index 10
# $Achievement_2 = Get-Content $clParametersPath | Select-Object -Index 11
# $Achievement_3 = Get-Content $clParametersPath | Select-Object -Index 12
# $Mission_Statement = Get-Content $clParametersPath | Select-Object -Index 13
# $Your_Name = Get-Content $clParametersPath | Select-Object -Index 14

# Read the content of the template document and store it in a variable
$Word = New-Object -ComObject word.application
$Document = $Word.Documents.Open($newClFullPath)

# Replace the placeholders with the input parameters
$parameterDictionary = @{
    "(Position_Name)" = $Position_Name;
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

