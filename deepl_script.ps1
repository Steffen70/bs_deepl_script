# DeepL API key
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
function Translate-Text($text) {
    $uri = "https://api-free.deepl.com/v2/translate"
    $body = @{
        auth_key    = $apiKey
        text        = $text
        target_lang = "EN"
    }
    $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body
    return $response.translations.text
}

# Iterate through each row
for ($row = 2; $row -le $rowMax; $row++) {
    $german = $worksheet.Cells.Item($row, 6).Text # Column D for German
    $english = $worksheet.Cells.Item($row, 7).Text # Column E for English

    if ([string]::IsNullOrWhiteSpace($english)) {
        $translation = Translate-Text $german
        $worksheet.Cells.Item($row, 7).Value = $translation # Update the English column with the translation
        Write-Output "Translated and updated row ${row}: $translation"
    }
}

# Save and close the workbook
$workbook.Close($true)
$excel.Quit()

# Run garbage collection - common practice when working with COM objects in PowerShell
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
