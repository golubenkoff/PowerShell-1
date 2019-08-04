function Set-DCOMLaunchPermissions ( [string] $appID, [string] $userOrSid, [string] $domain )
{
  $app = Get-WmiObject -Query ('SELECT * FROM Win32_DCOMApplicationSetting WHERE AppId = "{0}"' -f $appId) -EnableAllPrivileges

  $sdRes = $app.GetLaunchSecurityDescriptor()

  "Current launch descriptor:"
  $sdRes

  $sd = $sdRes.Descriptor

  "Creating trustee..."
  $trustee = ([wmiclass] 'Win32_Trustee').CreateInstance()
  
  if ($domain -eq 'nt authority')
  {
    $sid = [wmi] "\\.\root\cimv2:Win32_SID.SID='$userOrSid'"
    
    $sid
    
    $trustee.SID = $sid.BinaryRepresentation
    $trustee.SIDLength = $sid.SIDLength
    $trustee.SIDString = $userOrSid
    $trustee.Domain = $sid.ReferencedDomainName
    $trustee.Name = $sid.AccountName
  }
  
  else
  {
    $trustee.Domain = $domain
    $trustee.Name = $userOrSid
  }

  $trustee

  $fullControl = 31
  $localLaunchActivate = 11

  $ace = ([wmiclass] 'Win32_ACE').CreateInstance()
  $ace.AccessMask = $localLaunchActivate
  $ace.AceFlags = 0
  $ace.AceType = 0
  $ace.Trustee = $trustee

  [System.Management.ManagementBaseObject[]] $newDACL = $sd.DACL + @($ace)

  $sd.DACL = $newDACL

  $app.SetLaunchSecurityDescriptor($sd)
}


Set-DCOMLaunchPermissions '{000C101C-0000-0000-C000-000000000046}' 'S-1-5-20' "NT Authority"
Set-DCOMLaunchPermissions '{000C101C-0000-0000-C000-000000000046}' 'WSS_ADMIN_WPG' $null
Set-DCOMLaunchPermissions '{61738644-F196-11D0-9953-00C04FD919C1}' 'WSS_ADMIN_WPG' $null
