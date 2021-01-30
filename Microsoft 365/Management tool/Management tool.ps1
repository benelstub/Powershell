#### M365 Management Script  

Set-ExecutionPolicy RemoteSigned

Install-Module -Name AzureAD
Install-Module -Name MSOnline
Install-Module -Name ExchangeOnlineManagement

Import-Module AzureAD
Import-Module ExchangeOnlineManagement


#region User Management Functions

#region View user details
function Get-User {
    $UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
    Get-AzureADUser -ObjectId $UPN | Out-GridView
}
#endregion

#region Creation of a new user

function New-User {
$DisplayName = Read-Host "Enter Display Name"
$FirstName = Read-Host "Enter First Name"
$SurName = Read-Host "Enter Surname"
$UPN = Read-Host "Enter user UPN (eg j.smith@contoso.onmicrosoft.com)"
$MailAlias = Read-Host "Mail Alias (Appears before the '@' symbol)"

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
$LicensePicker = Get-AzureADSubscribedSku | Select-Object SkuPartNumber | Out-GridView -Title "Select the License you wish to apply" -OutputMode Single 
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

#endregion

#region User Mangement Menu
function Show-UsrMgmtMenu {
Do{
        Clear-Host
        Write-Host "
        ========User Management========
        "
        Write-Host "
        View user details : 1
        Create new user   : 2
        Change password   : 3
        Set access        : 4
        Add license       : 5
        Quit              : Q
        "
        $Opt = Read-Host
            Switch ($Opt)
            {

                1{Get-User}
                2{New-User}
                3{Set-Passowrd}
                4{Set-Access}
                5{Add-License}
                
            }
    }Until ($Opt -eq "q")
}
#endregion

#region User Reporting Functions

#region All users
function Get-AllUserReport {
    Get-AzureADUser -All True | Export-Csv C:\Temp\allusers.csv -NoTypeInformation 
}
#endregion

#region Enabled users
function Get-EnabledUserReport {
    Get-AzureADUser -Filter "AccountEnabled eq true" | Export-Csv C:\Temp\enabledusers.csv -NoTypeInformation 
}
#endregion

#region Disabled users
function Get-DisabledUserReport {
    Get-AzureADUser -Filter "AccountEnabled eq false" | Export-Csv C:\Temp\disabledusers.csv -NoTypeInformation 
}
#endregion

#region Recently created users
function Get-RecentUserCreated {
    Write-Host "In prog"
}
#endregion

#region Revently deleted users
function Get-RecentUserDeleted {
    Write-Host "In prog"
} 
#endregion


#endregion

#region User Reporting Menu
function Show-UsrReportMenu {
Do{
        Clear-Host
        Write-Host "
        ========User Reporting========
        "
        Write-Host "
        All users               : 1
        Enabled users           : 2
        Disabled users          : 3
        Recently created users  : 4
        Recently deleted users  : 5
        Quit                    : Q
        "
        $Opt = Read-Host
            Switch ($Opt)
            {

                1{Get-AllUserReport}
                2{Get-EnabledUserReport}
                3{Get-DisabledUserReport}
                4{Get-RecentUserCreated}
                5{Get-RecentUserDeleted}
                
            }
    }Until ($Opt -eq "q")
}
#endregion


#region EXO Management Menu
function Show-EXOMgmtMenu {
Do{
        Clear-Host
        Write-Host "
        ========EXO Management========
        "
        Write-Host "
        Set auto-reply               : 1
        Set delegate permissions     : 2
        Remove delegate permissions  : 3
        Set forwarding               : 4
        Remove forward               : 5
        Quit                         : Q
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

#region EXO Reporting Menu
function Show-EXOReportMenu {
Do{
        Clear-Host
        Write-Host "
        ========EXO Reporting========
        "
        Write-Host "
        All mailboxes                       : 1
        Recently created                    : 2
        Recently deleted                    : 3
        User mailbox delegate permissions   : 4
        Shared mailbox delegate permissions : 5
        Quit                                : Q
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

#region Change Customer
function Connect-Customer {
    $credential = Get-Credential
    
    Connect-AzureAD -Credential $credential
    Connect-ExchangeOnline -Credential $credential -ShowProgress $true
    Connect-MsolService -Credential $credential
}
#endregion

#region Tenant info
function Get-TenantInfo{
    Get-MsolCompanyInformation | Out-GridView
}
#endregion

#region Main Menu
function Show-HomeMenu {
Do{
        Clear-Host
        Write-Host "
        ========Main========
        "
        Write-Host "**You must select option 1 on launch**" -ForegroundColor Yellow
        Write-Host "
        Switch Customer     : 1
        List tenant details : 2
        User Management     : 3
        User Reporting      : 4
        EXO Management      : 5
        EXO Reporting       : 6
        Quit                : Q
        "
        $Opt = Read-Host
            Switch ($Opt)
            {

                1{Connect-Customer}
                2{Get-TenantInfo}
                3{Show-UsrMgmtMenu}
                4{Show-UsrReportMenu}
                5{Show-EXOMgmtMenu}
                6{Show-EXOReportMenu}
                
            }
    }Until ($Opt -eq "q")
}
#endregion

Show-HomeMenu