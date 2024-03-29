##################################################################################################################
#
# Microsoft Premier Field Engineering
# jesse.esquivel@microsoft.com
# MSResponderDoDConfig.ps1
# -updated 02/24/16
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

$scriptpath = $MyInvocation.MyCommand.Path

#Modify the following two values to match your environment
$OCSPSigningTemplate = "ProsewareAzure-OCSPResponseSigning" #Get the certificate template name from the certificate templates snap-in
#login to your MSFT issuing CA and issue the certutil -cainfo command to get the following values for the next line in this format:  <DNS NAME>\<CA Name>
$CACOnfig = "proazureca01.proseware.com\Proseware Azure Subordinate Certificate Authority"

function create-DoDRevocationProvider($DoDCert,$crl)
    {   
        $dir = Split-Path $scriptpath
        $file = Get-ChildItem "$dir\$DoDCert"
        $certname = $file.name
        $cert = $dir + "\" + $certname

        # Get the certificate from the local store by using it's DN
        $CaCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate
        $CaCert.Import($cert)
        $CaCert = $CaCert.GetRawCertData()

        # Save the desired OcspProperties in a collection object
        $OcspProperties = New-Object -com "CertAdm.OCSPPropertyCollection"
        $OcspProperties.CreateProperty("BaseCrlUrls", $crl)
        $OcspProperties.CreateProperty("RevocationErrorCode", 0)
        # Sets the refresh interval to 24 hours (time is specified in milliseconds)
        $OcspProperties.CreateProperty("RefreshTimeOut", 86400000)
        
        # Save the baseName in a variable, this is the filename without extension
        # eg. basename of certificate.cer is certificate
        $certBaseName = $file.BaseName

        # Save the current configuration in an OcspAdmin object
        $OcspAdmin = New-Object -com "CertAdm.OCSPAdmin"
        $OcspAdmin.GetConfiguration($env:computername, $true)

        # Create a new revocation configuration
        $NewConfig = $OcspAdmin.OCSPCAConfigurationCollection.CreateCAConfiguration($certBaseName, $CaCert)
        $NewConfig.HashAlgorithm = "SHA1"
        $NewConfig.SigningFlags = 0x294
        $NewConfig.CAConfig = $CAConfig
        $NewConfig.SigningCertificateTemplate = $OCSPSigningTemplate
        $NewConfig.ProviderProperties = $OcspProperties.GetAllProperties()
        $NewConfig.ProviderCLSID = "{4956d17f-88fd-4198-b287-1e6e65883b19}"
        $NewConfig.ReminderDuration = 90

        # Commit the new configuration to the server
        $OcspAdmin.SetConfiguration($env:computername, $true)
    }

#Call function to create the required revocation configurations

#existing configs
create-DoDRevocationProvider "DoD Root CA2.cer" "http://crl.disa.mil/getcrl?DoD%20Root%20CA%202"
create-DoDRevocationProvider "DoD Intermediate CA-1.cer" "http://crl.disa.mil//getcrl?DOD%20INTERMEDIATE%20CA-1"
create-DoDRevocationProvider "DoD Intermediate CA-2.cer" "http://crl.disa.mil//getcrl?DOD%20INTERMEDIATE%20CA-2"

<#expired
create-DoDRevocationProvider "DoD CA-19.cer" "http://crl.disa.mil//getcrl?DOD%20CA-19"
create-DoDRevocationProvider "DoD CA-20.cer" "http://crl.disa.mil//getcrl?DOD%20CA-20"
create-DoDRevocationProvider "DoD CA-21.cer" "http://crl.disa.mil//getcrl?DOD%20CA-21"
create-DoDRevocationProvider "DoD CA-22.cer" "http://crl.disa.mil//getcrl?DOD%20CA-22"
create-DoDRevocationProvider "DoD CA-23.cer" "http://crl.disa.mil//getcrl?DOD%20CA-23"
create-DoDRevocationProvider "DoD CA-24.cer" "http://crl.disa.mil//getcrl?DOD%20CA-24"
create-DoDRevocationProvider "DoD CA-25.cer" "http://crl.disa.mil//getcrl?DOD%20CA-25"
create-DoDRevocationProvider "DoD CA-26.cer" "http://crl.disa.mil//getcrl?DOD%20CA-26"
create-DoDRevocationProvider "DoD EMAIL CA-19.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-19"
create-DoDRevocationProvider "DoD EMAIL CA-20.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-20"
create-DoDRevocationProvider "DoD EMAIL CA-21.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-21"
create-DoDRevocationProvider "DoD EMAIL CA-22.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-22"
create-DoDRevocationProvider "DoD EMAIL CA-23.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-23"
create-DoDRevocationProvider "DoD EMAIL CA-24.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-24"
create-DoDRevocationProvider "DoD EMAIL CA-25.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-25"
create-DoDRevocationProvider "DoD EMAIL CA-26.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-26"
#>

create-DoDRevocationProvider "DoD CA-27.cer" "http://crl.disa.mil//getcrl?DOD%20CA-27"
create-DoDRevocationProvider "DoD CA-28.cer" "http://crl.disa.mil//getcrl?DOD%20CA-28"
create-DoDRevocationProvider "DoD CA-29.cer" "http://crl.disa.mil//getcrl?DOD%20CA-29"
create-DoDRevocationProvider "DoD CA-30.cer" "http://crl.disa.mil//getcrl?DOD%20CA-30"
create-DoDRevocationProvider "DoD CA-31.cer" "http://crl.disa.mil//getcrl?DOD%20CA-31"
create-DoDRevocationProvider "DoD CA-32.cer" "http://crl.disa.mil//getcrl?DOD%20CA-32"
create-DoDRevocationProvider "DoD EMAIL CA-27.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-27"
create-DoDRevocationProvider "DoD EMAIL CA-28.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-28"
create-DoDRevocationProvider "DoD EMAIL CA-29.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-29"
create-DoDRevocationProvider "DoD EMAIL CA-30.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-30"
create-DoDRevocationProvider "DoD EMAIL CA-31.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-31"
create-DoDRevocationProvider "DoD EMAIL CA-32.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-32"

#new configs - 02/24/16 for Bob!!
create-DoDRevocationProvider "DoD Root CA3.cer" "http://crl.disa.mil/getcrl?DoD%20Root%20CA%203"
create-DoDRevocationProvider "DoD EMAIL CA-33.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-33"
create-DoDRevocationProvider "DoD EMAIL CA-34.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-34"
create-DoDRevocationProvider "DoD EMAIL CA-39.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-39"
create-DoDRevocationProvider "DoD EMAIL CA-40.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-40"
create-DoDRevocationProvider "DoD EMAIL CA-41.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-41"
create-DoDRevocationProvider "DoD EMAIL CA-42.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-42"
create-DoDRevocationProvider "DoD EMAIL CA-43.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-43"
create-DoDRevocationProvider "DoD EMAIL CA-44.cer" "http://crl.disa.mil//getcrl?DOD%20EMAIL%20CA-44"
create-DoDRevocationProvider "DoD ID CA-33.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-33"
create-DoDRevocationProvider "DoD ID CA-34.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-34"
create-DoDRevocationProvider "DoD ID CA-39.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-39"
create-DoDRevocationProvider "DoD ID CA-40.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-40"
create-DoDRevocationProvider "DoD ID CA-41.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-41"
create-DoDRevocationProvider "DoD ID CA-42.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-42"
create-DoDRevocationProvider "DoD ID CA-43.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-43"
create-DoDRevocationProvider "DoD ID CA-44.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20CA-44"
create-DoDRevocationProvider "DoD ID SW CA-35.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20SW%20CA-35"
create-DoDRevocationProvider "DoD ID SW CA-36.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20SW%20CA-36"
create-DoDRevocationProvider "DoD ID SW CA-37.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20SW%20CA-37"
create-DoDRevocationProvider "DoD ID SW CA-38.cer" "http://crl.disa.mil//getcrl?DOD%20ID%20SW%20CA-38"