<#
.Synopsis
   Retrieve CRL and verify ThisUpdate and NextUpdate is satisfactory
.DESCRIPTION
   Uses certutil -dump <crlFile> to determine CRL date and verify it is greater than X minutes
.EXAMPLE
   Get-CrlExpiration -Uri "http://pki.example.com/pki/CA1.crl" -AlertThreshold 90
   Get-CrlExpiration -Uri "http://pki.example.com/pki/CA1+.crl" -AlertThreshold 90
#>
function Get-CrlExpiration
{
    [CmdletBinding()]
    [Alias()]
    param (
        # number of minutes remaining until NextUpdate for triggering alert
        [int]$AlertThreshold = 90,
        # URL to CRL
        [Parameter(Mandatory=$true)]
        [string]$Uri
    )

    Function Remove-InvalidFileNameChars {
        param(
        [Parameter(Mandatory=$true,
            Position=0,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String]$Name
        )

        $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
        $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
        return ($Name -replace $re)
    }

    try 
    {
        # download CRL file
        $uriFilename = Remove-InvalidFileNameChars ($Uri -split "\/")[-1]
        $outFile = "$env:TEMP\$uriFilename-$([DateTime]::Now.ToString("yyyyMMdd-HHmmss"))"
        Invoke-WebRequest -Uri $Uri -OutFile $outFile -UseBasicParsing
        
        # parse certutil output
        $crlDate = certutil -dump $outFile
        $thisUpdate = (($crlDate -match "ThisUpdate\:") -split "\: ")[1]
        $nextUpdate = (($crlDate -match "NextUpdate\:") -split "\: ")[1]

        # determine number of minutes remaining
        $timeLeft = New-TimeSpan -Start (Get-Date) -End (Get-Date $nextUpdate)

        # determine pass/fail
        If ($timeLeft.TotalMinutes -gt $AlertThreshold){
            $Healthy = $True   # pass
        } else {
            $Healthy = $False  # fail
        }

        # Output 
        Write-Verbose "Uri: $Uri. OutFile: $outFile. ThisUpdate: $thisUpdate. NextUpdate: $NextUpdate. TotalMinutes Remaining: $($timeLeft.TotalMinutes). AlertThreshold: $AlertThreshold. Healthy: $Healthy"
        New-Object -TypeName PSObject -Property ([ordered]@{
            Uri              = $Uri
            ThisUpdate       = $thisUpdate
            NextUpdate       = $nextUpdate
            TotalMinutesLeft = $($timeLeft.TotalMinutes)
            Healthy          = $Healthy
        })
    } 
    catch 
    {
        Write-Error $_
        continue
    }
}
