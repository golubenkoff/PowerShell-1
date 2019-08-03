<#
.Synopsis
   Retrieve the nearest requestID based on specified date on the target Certificate Authority using binary search (defaulting to low value if no match found)
.DESCRIPTION
   Requires PSPKI PowerShell module   
   Uses certutil -view -restrict "$RequestID=$" to retrieve Newest request from the CA
   Increments through all certificate requests on the target Certificate Authority until it finds the nearest request ID since the specified date.
.EXAMPLE
   Get-CA | Get-NearestCertificateRequest
.EXAMPLE
   Get-CertificateAuthority "CA3" | Get-NearestCertificateRequest
#>
function Get-NearestCertificateRequest
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    param (
        # requires PSPKI PowerShell module, creates a [PKI.CertificateServices.CertificateAuthority] object
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PKI.CertificateServices.CertificateAuthority[]]$CertificateAuthority,
        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )

    begin {

    }
    process {
        foreach ($CA in $CertificateAuthority)
        {
            try 
            {
                # parse certutil output for last certificate request issued from CA
                $configString = $CA.ConfigString              
                $newestRequestID = certutil -view -restrict "RequestId=$" -config $configString -out RequestID
                $newestRequestID = $newestRequestID -match "Issued Request ID\:"
                $newestRequestID = $newestRequestID -replace "Issued Request ID\:",""
                $newestRequestID = ($newestRequestID -split "\(")[1]
                $newestRequestID = $newestRequestID -replace "\)",""
                Write-Verbose "[$($CA.Name)] Found newest certificate request ID: $newestRequestID"

                # get information about latest request in database
                $newestRequest = Get-AdcsDatabaseRow -CertificationAuthority $CA -RowID $newestRequestID
                Write-Verbose "[$($CA.Name)] Found information about newest certificate request ID: $newestRequestID. Request.SubmittedWhen Date: $($newestRequest.'Request.SubmittedWhen')"

                # set the search value
                $SearchVal = $Date

                # set the current request and index
                $LowIndex = 0                              # Low side of array segment
                $Counter = 0
                $TempVal = ""                              # Used to determine end of search where $Found = $False
                $HighIndex = $newestRequestID              # High Side of array segment
                [int]$MidPoint = ($HighIndex-$LowIndex)/2  # Mid point of array segment
                $Found = $False

                While($LowIndex -le $HighIndex)
                {
                    # Determine the timestamp of the MidPoint
                    $InputRequest = (Get-AdcsDatabaseRow -CertificationAuthority $CA -RowID $MidPoint)
                    While ($InputRequest -eq $null)
                    {
                        # if database row is null, increment backwards from midpoint to find a valid row....
                        $MidPoint = $MidPoint-1
                        $InputRequest = (Get-AdcsDatabaseRow -CertificationAuthority $CA -RowID $MidPoint)
                    }
                    $MidVal = $InputRequest.'Request.SubmittedWhen' 
                        
                    # If identical, the search has completed and $Found = $False                                                                    
                    If($TempVal -eq $MidVal)
                    {
                        # If low index is 0, set it to 1 (there is no ADCS database row for 0)
                        If ($LowIndex -eq 0)                            
                        {                            
                            $LowIndex = 1
                        }
                        
                        # Output low index as final index
                        $Found = $False                        
                        $OutputRequest = (Get-AdcsDatabaseRow -CertificationAuthority $CA -RowID $LowIndex)
                        Write-Verbose "[$($CA.Name)] Binary Search has completed - exact match not found in $Counter passes. Index set to Low value: $LowIndex. Request.SubmittedWhen Date: $($OutputRequest.'Request.SubmittedWhen')"                        
                        $LowIndex
                        Return
                    }
                    else
                    {
                        #Update the TempVal. Search continues.
                        $TempVal = $MidVal
                    }

                    # write-host "Midval is: $midval"
                    # Write-host "Low is $lowindex, Mid is $midpoint, High is $HighIndex"
                    Write-Verbose "[$($CA.Name)] Binary Search ongoing... Search date: $SearchVal. Current date: $MidVal. LowIndex is $LowIndex, MidPoint is $MidPoint, HighIndex is $HighIndex."
                    If ($SearchVal -lt $MidVal)
                    {
                        # Write-Host "SV < MV"
                        $Counter++
                        $HighIndex = $MidPoint 
                        [int]$MidPoint = (($HighIndex-$LowIndex)/2 +$LowIndex)
                    }
                    If($SearchVal -gt $MidVal)
                    {                        
                        # Write-Host "SV > MV"
                        $Counter++
                        $LowIndex = $MidPoint 
                        [int]$MidPoint = ($MidPoint+(($HighIndex - $MidPoint) / 2))         
                    }
                    If($SearchVal -eq $MidVal)
                    {
                        # Output midpoint as final index
                        $Found = $True
                        $OutputRequest = (Get-AdcsDatabaseRow -CertificationAuthority $CA -RowID $MidPoint)
                        Write-Verbose "[$($CA.Name)] Binary Search has completed - exact match found in $Counter passes. Index set to exact value: $MidPoint. Request.SubmittedWhen Date: $($OutputRequest.'Request.SubmittedWhen')"
                        $MidPoint
                        Return
                    }
                } #End
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
