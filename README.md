# Excel Translation Tool

This tool can be used to translate SPi texts from german to english.

## Installation

The script uses Office 365 COM to interact with Excel. Make sure you have the required components installed. (Office 365, Excel)

## Usage

**PowerShell 7**

```shell
pwsh .\deepl_script.ps1
```

**Windows PowerShell**

```shell
powershell .\deepl_script.ps1
```

## Adittional Tips

Set the execution policy to RemoteSigned to run the script.

```shell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```
