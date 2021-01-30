Set-ExecutionPolicy RemoteSigned
Install-Module -Name AzureAD
Add-Type -AssemblyName 'System.Web'

$length = 10
$nonAlphaChars = 4

$credential = Get-Credential
Connect-AzureAD -Credential $credential

$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
$newPassword = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
$secPassword = ConvertTo-SecureString $newPassword -AsPlainText -Force

Set-AzureADUserPassword -ObjectId $UPN -Password $secPassword 

Write-Host $newPassword


