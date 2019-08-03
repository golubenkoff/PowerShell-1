<#
.Synopsis
   Retrieve newest requestID received on the target Certificate Authority
.DESCRIPTION
   Requires PSPKI PowerShell module
   
   Uses certutil -view -restrict "$RequestID=$" to retrieve Newest request from the CA
   
   https://blogs.technet.microsoft.com/pki/2008/10/03/disposition-values-for-certutil-view-restrict-and-some-creative-samples/
.EXAMPLE
   Get-CA | Get-NewestCertificateRequest
.EXAMPLE
   Get-CertificateAuthority "CA3" | Get-NewestCertificateRequest
#>
function Get-NewestCertificateRequest
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
                $newestRequestID = certutil -view -restrict "RequestId=$" -config $configString -out RequestID
                $newestRequestID = $newestRequestID -match "Issued Request ID\:"
                $newestRequestID = $newestRequestID -replace "Issued Request ID\:",""
                $newestRequestID = ($newestRequestID -split "\(")[1]
                $newestRequestID = $newestRequestID -replace "\)",""
                Write-Verbose "Found RequestID of newest certificate request on CA: $newestRequestID"

                # output
                $newestRequestID
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
