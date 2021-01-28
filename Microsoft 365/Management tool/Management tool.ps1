#### User management on M365  

##You must install the Azure Active Directory Module and EXO V2 Module. See below:
Install-Module -Name AzureAD
#Install-Module -Name ExchangeOnlineManagement

Set-ExecutionPolicy RemoteSigned

#region Change Customer
function Connect-Customer {
$credential = Get-Credential
#Import-Module ExchangeOnlineManagement

Connect-AzureAD -Credential $credential
#Connect-ExchangeOnline -Credential $credential -ShowProgress $true

}
#endregion

#region Creation of a new user

function New-User {
$DisplayName = Read-Host "Enter Display Name"
$FirstName = Read-Host "Enter First Name"
$SurName = Read-Host "Enter Surname"
$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
$MailAlias = Read-Host "Mail Alias (Appears before the '@' symbol"

$PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password= Read-Host "Enter a Password"

New-AzureADUser -DisplayName $DisplayName -GivenName $FirstName -SurName $SurName -UserPrincipalName $UPN -UsageLocation GB -MailNickName $MailAlias -PasswordProfile $PasswordProfile -AccountEnabled $true
}

#endregion

#region Change Password

function Set-Password {
$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
$newPassword = Read-Host "Enter a new password"
$secPassword = ConvertTo-SecureString $newPassword -AsPlainText -Force

Set-AzureADUserPassword -ObjectId $UPN -Password $secPassword 
}

#endregion

#region Assign Sign-In Rights

function Set-Access {
$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"

    function EnableAccess {
    Set-AzureADUser -ObjectID $UPN -AccountEnabled $true
    }

    function DisableAccess {
    Set-AzureADUser -ObjectID $UPN -AccountEnabled $false
    }
    
    Do{
        Write-Host "
        ========User Sign-in========
        "
        Write-Host "
        Enable : 1
        Disable: 2
        "
        $Opt = Read-Host
            Switch ($Opt)
            {

                1{
                Write-Host "Enabled Sign-in" -ForegroundColor Green
                EnableAccess}
                2{
                Write-Host "Disabled Sign-in" -ForegroundColor Red
                DisableAccess}
                
            }
    }Until ($Opt -eq '1' -or $Opt -eq '2')
    Write-Host "Please wait..."
    Start-Sleep -s 3.5
}

#endregion

#region Add a license

function Add-License {
$LicensePicker = Get-AzureADSubscribedSku | Select SkuPartNumber | Out-GridView -Title "Select the License you wish to apply" -OutputMode Single 
$LicenseSelect = $LicensePicker.SkuPartNumber
    Write-Host "You selected:"$LicenseSelect

$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
    Set-AzureADUser -ObjectID $UPN -UsageLocation GB

$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $LicenseSelect -EQ).SkuID
$LicenseToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$LicenseToAssign.AddLicenses = $License

Set-AzureADUserLicense -ObjectId $UPN -AssignedLicenses $LicenseToAssign

}

#endregion

#region Main Menu
function Show-Menu {
Do{
        Clear-Host
        Write-Host "
        ========Main========
        "
        Write-Host "**You must select option 1 on launch**" -ForegroundColor Yellow
        Write-Host "
        Switch Customer : 1
        Create User     : 2
        Change password : 3
        Set access      : 4
        Add license     : 5
        Quit            : Q
        "
        $Opt = Read-Host
            Switch ($Opt)
            {

                1{Connect-Customer}
                2{New-User}
                3{Set-Passowrd}
                4{Set-Access}
                5{Add-License}
                
            }
    }Until ($Opt -eq "q")
}
#endregion

Show-Menu