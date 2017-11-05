#Requires -Modules ActiveDirectory
###############################################################################
#                                                                             #
#            Active Directory security group membership audit                 #
#     https://technet.microsoft.com/en-us/library/dn579255(v=ws.11).aspx      #
#                                                                             #
# Default groups, such as the Domain Admins group, are security groups that   #
#  are created automatically when you create an Active Directory domain. You  #
#  can use these predefined groups to help control access to shared resources #
#  and to delegate specific domain-wide administrative roles. Many default    #
#  groups are automatically assigned a set of user rights that authorize      #
#  members of the group to perform specific actions in a domain, such as      #
#  logging on to a local system or backing up files and folders. For example, #
#  a member of the Backup Operators group has the right to perform backup     #
#  operations for all domain controllers in the domain.                       #
#                                                                             #
###############################################################################
#                                                                             #
#  date:   2017-11-04                                                         #     
#  note:   made with love by github.com/milesgratz                            #
#                                                                             #
###############################################################################

# Define folder to save results (e.g "C:\Temp")
# if null, saves to $env:USERPROFILE\Desktop
$savePath = 

# Create an empty array for results + define list of groups
$Results = @()
$Groups = (
    'Access Control Assistance Operators',
    'Account Operators',
    'Administrators',
    #'Allowed RODC Password Replication Group',
    'Backup Operators',
    'Certificate Service DCOM Access',
    'Cert Publishers',
    #'Cloneable Domain Controllers',
    'Cryptographic Operators',
    #'Denied RODC Password Replication Group',
    'Distributed COM Users',
    #'DnsUpdateProxy',
    'DnsAdmins',
    'Domain Admins',
    #'Domain Computers',
    #'Domain Controllers',
    #'Domain Guests',
    #'Domain Users',
    'Enterprise Admins',
    #'Enterprise Read-Only Domain Controllers',
    'Event Log Readers',
    'Group Policy Creator Owners',
    #'Guests',
    'Hyper-V Administrators',
    'IIS_IUSRS',
    'Incoming Forest Trust Builders',
    'Network Configuration Operators',
    'Performance Log Users',
    'Performance Monitor Users',
    #'Pre-Windows 2000 Compatible Access',
    'Print Operators',
    'Protected Users',
    #'RAS and IAS Servers',
    #'RDS Endpoint Servers',
    #'RDS Management Servers',
    #'RDS Remote Access Servers',
    #'Read-Only Domain Controllers',
    'Remote Desktop Users',
    'Remote Management Users',
    #'Replicator',
    'Schema Admins',
    'Server Operators',
    #'Terminal Server License Servers',
    #'Users',
    'Windows Authorization Access Group',
    'WinRMRemoteWMIUsers'
)

# Loop through groups and adding users 
$Groups | ForEach-Object {
    $GroupName = $_
    Write-Host "[$GroupName] Finding members in group"
    $Results += Get-ADGroupMember $GroupName -Recursive | ForEach-Object { 
        If ($_.objectClass -eq 'user')
        { 
            Write-Host "[$GroupName] Adding $($_.Name) to results" -ForegroundColor Green
            Get-ADUser $_ -Properties * | Select @{n='Group';e={"$GroupName"}},*
        }
        Else 
        {
            Write-Host "[$GroupName] Skipping $($_.objectClass) object: $($_.Name)" -ForegroundColor Yellow
        }
    }
}

# Let's trim out some attributes
$Output = $Results | Select `
    Group,
    Name,
    Enabled,
    Description,
    Created,
    LastLogonDate,
    PasswordLastSet,
    LockedOut,
    AccountLockoutTime,
    LastBadPasswordAttempt,
    badPwdCount,
    logonCount,
    PasswordExpired,
    PasswordNeverExpires,
    AccountNotDelegated,
    SmartcardLogonRequired

# Define output filename
$dateFormat = Get-Date -Format yyyy-MM-dd
If ($savePath){ $filePath = "$savePath\AD Security Group Audit $dateFormat.csv" }
Else { $filePath = "$env:USERPROFILE\Desktop\AD Security Group Audit $dateFormat.csv" } 

# Saving results to file
Write-Host "Script completed. Saving results to $filePath" -ForegroundColor Green
$Output | Export-Csv $filePath -NoTypeInformation