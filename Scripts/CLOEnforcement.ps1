##################################################################################################################
#
# Microsoft Premier Field Engineering
# jesse.esquivel@microsoft.com
# CLOEnforcement.ps1
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

Import-Module ActiveDirectory

#query AD for user accounts using the following LDAP filter (users that are not currently smart card enforced
$OUQuery1 = Get-ADObject -LDAPFilter "(&(ObjectCategory=Person)(ObjectClass=User)(!userAccountControl:1.2.840.113556.1.4.803:=262144))" -Properties distinguishedName,useraccountcontrol -server localhost:389 -searchBase "OU=Enforcement Test,DC=somedomain,DC=com" -searchScope subtree
$OUQuery2 = Get-ADObject -LDAPFilter "(&(ObjectCategory=Person)(ObjectClass=User)(!userAccountControl:1.2.840.113556.1.4.803:=262144))" -Properties distinguishedName,useraccountcontrol -server localhost:389 -searchBase "OU=Infrastructure,DC=somedomain,DC=com" -searchScope subtree
$OUQuery3 = Get-ADObject -LDAPFilter "(&(ObjectCategory=Person)(ObjectClass=User)(!userAccountControl:1.2.840.113556.1.4.803:=262144))" -Properties distinguishedName,useraccountcontrol -server localhost:389 -searchBase "OU=Development,DC=somedomain,DC=com" -searchScope subtree

function Enforce-Users($Users)
{
    #here we test the $users  variable to ensure it isn't NULL"
    If ($users)
    {
        #do work here, iterate through the $users variable (Get-ADObject query return set)
        foreach($user in $users)
        {
            $exception = "False" 
            #test for a clo exception group
            $userGroupMemberships = get-ADPrincipalGroupMembership -server localhost:389 -identity $user.distinguishedName
            foreach($groupmembership in $userGroupMemberships)
                {
                    If($groupmembership -like "*CLO_Exceptions*")
                        {
                            #write-host "Not Enforced: " $user.name
                            $exception = "True"
                        }
                }
             if($exception -eq "False")
                {
                    #not an exception, enforce smart card required for interactive logon
                    $user.userAccountControl = $user.userAccountControl -bor 262144
                    Set-ADObject -Instance $user 
                    #write-host "Enforced: " $user.name
                }  
        }
    }
}

Enforce-Users($OUQuery1)
Enforce-Users($OUQuery2)
Enforce-Users($OUQuery3)