# DeepL API key - this api key is invalid, replace it with your own
$apiKey = "1e6a6f8e-949f-4adc-9eb2-e1c5f59eb094:fx"

# Input file
$inputFile = "input.xlsx"

# Create a new Excel application instance
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false # Set this to $true if you want to see the Excel application

# Open the workbook
$workbook = $excel.Workbooks.Open((Get-Location).Path + "/" + $inputFile)

# Select the first worksheet
$worksheet = $workbook.Worksheets.Item(1)

# Calculate the used range
$usedRange = $worksheet.UsedRange
$rowMax = $usedRange.Rows.Count

# Debugging - with a fixed row count to avoid unnecessary API calls
#$rowMax = 10

# Function to translate text using DeepL
function Translate-Text {
    param (
        [string]$text,
        [string]$targetLang = "EN"
    )
    $uri = "https://api-free.deepl.com/v2/translate"
    $body = @{
        auth_key    = $apiKey
        text        = $text
        target_lang = $targetLang
    }
    $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body
    return $response.translations.text
}

$germanColumn = 1; # Column D for German
$englishColumn = 7; # Column E for English
# $frechColumn = 3; # Column F for French
# $italianColumn = 5; # Column G for Italian

# Iterate through each row
for ($row = 2; $row -le $rowMax; $row++) {
    # Check if the colum number matches the column number in your Excel file
    $german = $worksheet.Cells.Item($row, $germanColumn).Text 
    $english = $worksheet.Cells.Item($row, $englishColumn).Text 

    if ([string]::IsNullOrWhiteSpace($english)) {
        $englishTranslation = Translate-Text $german
        $worksheet.Cells.Item($row, $englishColumn).Value = $englishTranslation # Update the English column with the translation
        Write-Output "Translated and updated row ${row}: $englishTranslation"
    }

    # You can add more languages if need to translate the text entries to more languages

    # $french = $worksheet.Cells.Item($row, $frechColumn).Text
    # $italian = $worksheet.Cells.Item($row, $italianColumn).Text
    # if ([string]::IsNullOrWhiteSpace($french)) {
    #     $frenchTranslation = Translate-Text $german -targetLang "FR"
    #     $worksheet.Cells.Item($row, $frechColumn).Value = $frenchTranslation # Update the French column with the translation
    #     Write-Output "Translated and updated row ${row}: $frenchTranslation"
    # }

    # if ([string]::IsNullOrWhiteSpace($italian)) {
    #     $italianTranslation = Translate-Text $german -targetLang "IT"
    #     $worksheet.Cells.Item($row, $italianColumn).Value = $italianTranslation # Update the Italian column with the translation
    #     Write-Output "Translated and updated row ${row}: $italianTranslation"
    # }
}

# Save and close the workbook
$workbook.Close($true)
$excel.Quit()

# Run garbage collection - common practice when working with COM objects in PowerShell
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
