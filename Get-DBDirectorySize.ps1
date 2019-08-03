<#
.Synopsis
   Retrieve newest requestID received on the target Certificate Authority
.DESCRIPTION
   Requires PSPKI PowerShell module
   
   Uses certutil -getreg config\DBDirectory to find database directory from the CA
   Uses Get-ChildItem to calculate size of database .edb file
   Returns integer object of database size in GB (rounded to 1 decimal places)
.EXAMPLE
   Get-CA | Get-DBDirectorySize
.EXAMPLE
   Get-CertificateAuthority "CA3" | Get-DBDirectorySize
#>
function Get-DBDirectorySize
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    param (
        # requires PSPKI PowerShell module, creates a [PKI.CertificateServices.CertificateAuthority] object
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PKI.CertificateServices.CertificateAuthority[]]$CertificateAuthority
    )

    begin {
           
    }
    process {
        foreach ($CA in $CertificateAuthority)
        {
            try 
            {
                # parse certutil output
                $configString = $CA.ConfigString
                $DBDirectory = certutil -config $configString  -getreg config\DBDirectory 
                $DBDirectory = $DBDirectory | Where-Object { $_ -match "DBDirectory\ REG_SZ" }
                $DBDirectory = ($DBDirectory -split "\= ")[1]

                # convert local path to remote path
                $localPath = $DBDirectory
                $remotePath = "\\" + $CA.ComputerName + "\" + $($DBDirectory -replace "\:","$")                

                # get newest file
                $DBFile = Get-ChildItem -Path "$remotePath" -Filter "*.edb" | Sort-Object $_.LastWriteTime -Descending | Select -First 1
                $DBFileSize = $DBFile.Length

                # output friendly size
                [System.Math]::Round((($DBFileSize)/1GB),1)
            }
            catch 
            {
                Write-Error $_
                continue
            }
        }
    }
    end {

    }
}
