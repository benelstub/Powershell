Set-ExecutionPolicy RemoteSigned
Import-Module ExchangeOnlineManagement

$Cred = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic –AllowRedirection

Import-PSSession $Session –DisableNameChecking

Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | 
Get-MailboxPermission |Select-Object Identity,User,AccessRights | 
Where-Object {($_.user -like '*@*')}|
Export-Csv C:\Temp\sharedfolders.csv  -NoTypeInformation 