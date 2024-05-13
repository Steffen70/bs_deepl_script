# Excel Translation Tool

This tool can be used to translate SPi texts from german to english.

## Installation

The script uses Office 365 COM to interact with Excel. Make sure you have the required components installed. (Office 365, Excel)

## Usage

**API Key**

You need to get an API key from [DeepL](https://www.deepl.com/pro-api) to use the translation service.

It should be sufficient to use the free plan to translate a view text entries.

**Excel File**

Replace the `input.xlsx` file with your own file. Keep in mind this script will overwrite the file with the translated text. So make a backup if you need to.

**Check the columns**

Please make sure the column numbers are correct in the script. I only translated the text entries to english.
But you may have additional columns for Italian and French translations.

**PowerShell 7**

```shell
pwsh .\deepl_script.ps1
```

**Note:** You may need to confirm the execution of the script by pressing `R` and then `Enter`.

## Adittional Tips

Set the execution policy to RemoteSigned to run the script.

```shell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
```
