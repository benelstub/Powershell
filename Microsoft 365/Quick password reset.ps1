Install-Module -Name AzureAD
Set-ExecutionPolicy RemoteSigned
$credential = Get-Credential
Connect-AzureAD -Credential $credential

$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
$newPassword = Read-Host "Enter a new password"
$secPassword = ConvertTo-SecureString $newPassword -AsPlainText -Force

Set-AzureADUserPassword -ObjectId $UPN -Password $secPassword 

Start-Sleep 3