##################################################################################################################
#
# Microsoft Premier Field Engineering
# jesse.esquivel@microsoft.com
# DoDSmartCardEnablement.ps1
# v1.0 Initial creation 8/2/13
# -Pull Principal name from smart card and write to UPN attribute of account.
# -Target User account objects need the following ACES (best to set on OU): 
#   SELF - Read public information
#   SELF - Write Public Information
#   SELF - Read employeeID
#   SELF - Write employeeID
#
# Microsoft Disclaimer for custom scripts
# ========================================
# The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. 
# Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
# The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone
# else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of
# business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts
# or documentation, even if Microsoft has been advised of the possibility of such damages.
# ========================================
#
##################################################################################################################

Function Get-EDIPI($userSN)
{
    $sn = $userSN.ToUpper()    
    $foundCert = "False"    
    $certStore = get-childitem -path cert:\CurrentUser\My
    $initStoreCount = $certStore.Count
    If($certStore)
    {
        ForEach($cert in $CertStore)
        {
            If($cert.Issuer -like "*DOD EMAIL*" -and $cert.Subject -like "*$sn*")
            { 
                If(($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "Subject Alternative Name" -and $cert.EnhancedKeyUsageList -like "*Smart Card Logon*"}))
                {
                    #write-host "Subject: " $cert.subject
                    #write-host "Issuer: " $cert.issuer
                    #write-host "Serial Number: " $cert.serialNumber
                    #write-host "friendly Name: " $cert.friendlyname
                    #write-host "Thumbprint: " $cert.Thumbprint
                    #write-host "key Size: " $cert.PrivateKey.KeySize
                    $san = ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1)
                    $arEDIPI = $san.split("=")
                    $strEDIPI = $arEDIPI[2].Substring(0,10)
                    $strCardUPN = $arEDIPI[2].Substring(0,14)
                    #Write-Host $strEDIPI
                    #Write-Host $strCardUPN
                    $foundCert = "True"
                    Break
                }
            }
        }
    }
    If($foundCert = "False")
    {
        $boxObject2 = New-Object -ComObject wscript.shell
        $msgBox2 = $boxObject2.popup("This script will run at every logon until you register your smart card.  Please Insert your CAC and click OK.",0,"Insert Smart Card")
        $timeout = New-TimeSpan -seconds 10        
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
        $objNotifyIcon.Icon = "C:\Program Files (x86)\Microsoft Office\Office14\FORMS\1033\SECRECL.ico"
        $objNotifyIcon.BalloonTipIcon = "Warning"
        $objNotifyIcon.BalloonTipText = "Attempting to read smart card..."
        $objNotifyIcon.BalloonTipTitle = "Active Directory Smart Card Enablement"
        $objNotifyIcon.Visible = $True
        $objNotifyIcon.ShowBalloonTip(9000) #set a few seconds less than $timeout
        $sw = [diagnostics.stopwatch]::StartNew()
        Do
        {
            $certStore = get-childitem -path cert:\CurrentUser\My
            $certStore.dispose
            #Write-Host "Reading Smart Card..."
            #write-host $sw.Elapsed
                If($sw.Elapsed -gt $timeout)
                {
                    Break
                }
                ElseIf($certStore.Count -gt $initStoreCount)
                {
                    Break
                }
        }
        While($foundCert = "False")
        $objNotifyIcon.Dispose()
        ForEach($cert in $CertStore)
        {
            If($cert.Issuer -like "*DOD EMAIL*" -and $cert.Subject -like "*$sn*")
            { 
                If(($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "Subject Alternative Name" -and $cert.EnhancedKeyUsageList -like "*Smart Card Logon*"}))
                {
                    #write-host "Subject: " $cert.subject
                    #write-host "Issuer: " $cert.issuer
                    #write-host "Serial Number: " $cert.serialNumber
                    #write-host "friendly Name: " $cert.friendlyname
                    #write-host "Thumbprint: " $cert.Thumbprint
                    #write-host "key Size: " $cert.PrivateKey.KeySize
                    #write-host "SubjectAltName: " ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1)
                    $san = ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1)
                    $arEDIPI = $san.split("=")
                    $strEDIPI = $arEDIPI[2].Substring(0,10)
                    $strCardUPN = $arEDIPI[2].Substring(0,14)
                    #Write-Host $strEDIPI
                    #Write-Host $strCardUPN
                    $foundCert = "True"
                    Break
                }
            }
        }
    }
    If($foundCert -eq "False")
    {
        $boxObject5 = New-Object -ComObject wscript.shell
        $msgBox5 = $boxObject5.popup("Logged on User: " + $strDomainUserName + $VBCrLf + $VBCrLf + "No suitable certificates were found." + `
        " Please contact your system administrator.",0,"Active Directory Smart Card Enablement")
        Exit
    }
}

Function Check-ADupn($CardUPN)
{
    #ADSI search filter, modify accordingly
    $strFilter = "(&(ObjectCategory=User)(userPrincipalName=$CardUPN))"

    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = $strFilter
    $result = $objSearcher.FindAll()

    If($result -ne $null)
    {
        $boxObject3 = New-Object -ComObject wscript.shell
        $msgBox3 = $boxObject3.popup("The UPN on your smart card is already in use in Active Directory, please contact your system administrator.",0,`
        "Active Directory UPN In Use")
        Exit
    }
}

#Main 
$VBCrLf = "`r`n"
$strName = $env:username
$strDomainUserName = [Security.Principal.WindowsIdentity]::GetCurrent().Name

#ADSI search filter, modify accordingly
$strFilter = "(&(ObjectCategory=User)(SAMAccountName=$strName))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$result = $objSearcher.FindAll()

$strADemployeeID = $result.GetDirectoryEntry().employeeid
$strADupn = $result.GetDirectoryEntry().userPrincipalName
$strADSurname = $result.GetDirectoryEntry().sn
$strADDN = $result.GetDirectoryEntry().distinguishedName


#check if employeeID attribute is populated, if so exit the script - the user has already been provisioned.
If($strADemployeeID -ne $null)
{
    #Write-Host "EmployeeID is populated! exit the script"
    Exit
}

$boxObject = New-Object -ComObject wscript.shell
$msgBox1 = $boxObject.popup("Logged on User: " + $strDomainUserName + $VBCrLf + $VBCrLf + "All interactive logons are required to be smart card enabled." + `
" Registration of your CAC digital signature certificates' Subject Alternative Name-Principal Name extension " + `
"is required for smart card logon.  You will need your CAC to register your Subject Alternative Name-Principal Name." + $VBCrLf`
+ $VBCrLf + "Click OK to continue.",0,"Active Directory Smart Card Enablement")

#Call function to get EDIPI and Principal Name from smart card
. Get-EDIPI($strADSurname)

#Call function to check if the UPN on the card is already in use in AD
Check-ADupn($strCardUPN)

#Write UPN and employeeID attributes to user account to provision for smart card logon
$objUser = [ADSI]"LDAP://$strADDN"
$objUser.Put("userPrincipalName", $strCardUPN)
$objUser.Put("employeeID", $strEDIPI)
$objUser.SetInfo()

$msgBox4 = $boxObject.popup("Logged on User: " + $strDomainUserName + $VBCrLf + $VBCrLf + "Your Active Directory Account has been updated with the following " + `
"UPN: " + $strCardUPN + $VBCrLf,0,"Active Directory Smart Card Enablement Completed")