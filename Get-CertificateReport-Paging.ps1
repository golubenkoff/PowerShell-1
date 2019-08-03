function Get-CertificateReport-Paging { 

    param (
        [CmdletBinding()]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PKI.CertificateServices.CertificateAuthority]$CertificateAuthority,
        [string]$CertificateTemplateOID,
        [datetime]$NotBeforeDate = (Get-Date "1970-01-01"),
        [string]$NotBeforeRequestID = 0,
        [int]$PageSize = 100,

        [Parameter(ParameterSetName='Issued')]
        [switch]$Issued,

        [Parameter(ParameterSetName='Revoked')]
        [switch]$Revoked,

        [Parameter(ParameterSetName='Pending')]
        [switch]$Pending,

        [Parameter(ParameterSetName='Failed')]
        [switch]$Failed
    )

    #======================================================================
    # Variables to customize for your environment
    #======================================================================
    $logPath = "$PSScriptRoot\Logs\Get-CertificateReport-Paging.log" 

    #======================================================================
    # Define logging function
    #======================================================================
    Function Write-Log {
        param(
            $Message,
            $LogPath,
            $Color = "Yellow",
            $Encoding = "Unicode"
        )
        Write-Host "$Message" -ForegroundColor $Color
        If ($LogPath -ne $null)
        {
            $logTime = Get-Date -format "yyyy-MM-dd HH:mm:ss zzz"
            Write-Output "[$logTime][$env:COMPUTERNAME] $Message" | Out-File -Append $LogPath -Encoding $Encoding
        }
    }

    #======================================================================
    # Creating new Event source on server
    #======================================================================
    try 
    {
        $eventSource = 'TestPKISource'
        if (!([System.Diagnostics.EventLog]::SourceExists($eventSource)))
        {
            New-EventLog -LogName Application -Source $eventSource -ErrorAction Stop
        }
    }
    catch
    {
        $errorMsg = $_.Exception.Message
        Write-Log "Failed to create new event log: $errorMsg" -LogPath $logPath
        Exit
    }

    #======================================================================
    # Import PowerShell PKI module
    #======================================================================
    Import-Module PSPKI -ErrorAction SilentlyContinue

    try 
    {
        # Verify PSPKI module is installed
        if (!(Get-Module PSPKI -ErrorAction Stop))
        { 
            Install-Module PSPKI -ErrorAction Stop -Force
            Import-Module PSPKI -ErrorAction Stop
        }
        else
        {
            # Do nothing - module already imported 
        }

        # Logging
        $outcome = "Successfully imported PowerShell PKI module"
        Write-Log $outcome -LogPath $logPath
    }
    catch
    {
        # Logging
        $errorMsg = $_.Exception.Message
        $outcome = "Failed to import PowerShell PKI module. Exception: $errorMsg"
        Write-Log $outcome -LogPath $logPath
        Write-EventLog -Message $outcome -LogName Application -Source $eventSource -EntryType Error -EventId 9000
        Exit
    }

    #======================================================================
    # Get issued certs
    #======================================================================
    if (!$Revoked -and !$Pending -and !$Failed)
    {
        try 
        {
            Write-Log "[$($CA.Name)] Searching for issued certificates" -LogPath $logPath

            $DateFilter = "NotBefore -ge $NotBeforeDate"
            $RequestIDFilter = "RequestID -ge $NotBeforeRequestID"
            $PageIndex = 1

            # perform paging
            If ($CertificateTemplateOID)
            {                               
                do
                {
                    # return all results of current page
                    $TemplateFilter = "CertificateTemplate -eq $CertificateTemplateOID"
                    $Paging = Get-IssuedRequest -CertificationAuthority $CertificateAuthority -Filter $TemplateFilter,$DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {
                        Write-Verbose "[$($CA.Name)] Processing issued request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Issued' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
            Else
            {
                do
                {
                    # return all results of current page
                    $Paging = Get-IssuedRequest -CertificationAuthority $CertificateAuthority -Filter $DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {
                        Write-Verbose "[$($CA.Name)] Processing issued request ID: $($_.RequestID)"                                     
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Issued' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
        }
        catch
        {
            $errorMsg = $_.Exception.Message
            $outcome = "Failed to query issued requests from CA. Exception: $errorMsg"
            Write-Log $outcome -LogPath $logPath
            Write-EventLog -Message $outcome -LogName Application -Source $eventSource -EntryType Error -EventId 9000
            Exit
        }
    }

    #======================================================================
    # Get revoked certs
    #======================================================================
    if ($Revoked)
    {
        Write-Log "[$($CA.Name)] Searching for revoked certificates" -LogPath $logPath
     
        try 
        {
            $DateFilter = "Request.RevokedWhen -ge $NotBeforeDate"
            $RequestIDFilter = "RequestID -ge $NotBeforeRequestID"
            $PageIndex = 1

            If ($CertificateTemplateOID)
            {                                
                do
                {
                    $TemplateFilter = "CertificateTemplate -eq $CertificateTemplateOID"
                    $Paging = Get-RevokedRequest -CertificationAuthority $CertificateAuthority -Filter $TemplateFilter,$DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {
                        Write-Verbose "[$($CA.Name)] Processing revoked request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Revoked' }}                   
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
            Else
            {
                do 
                {
                    $Paging = Get-RevokedRequest -CertificationAuthority $CertificateAuthority -Filter $DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {
                        Write-Verbose "[$($CA.Name)] Processing revoked request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Revoked' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
        }
        catch
        {
            $errorMsg = $_.Exception.Message
            $outcome = "Failed to query revoked requests from CA. Exception: $errorMsg"
            Write-Log $outcome -LogPath $logPath
            Write-EventLog -Message $outcome -LogName Application -Source $eventSource -EntryType Error -EventId 9000
            Exit
        }
    }

    #======================================================================
    # Get pending certs
    #======================================================================
    if ($Pending)
    {
        Write-Log "[$($CA.Name)] Searching for pending certificate requests" -LogPath $logPath

        try 
        {
            $DateFilter = "Request.SubmittedWhen -ge $NotBeforeDate"
            $RequestIDFilter = "RequestID -ge $NotBeforeRequestID"
            $PageIndex = 1

            If ($CertificateTemplateOID)
            {
                do
                {
                    $TemplateFilter = "CertificateTemplate -eq $CertificateTemplateOID"
                    $Paging = Get-PendingRequest -CertificationAuthority $CertificateAuthority -Filter $TemplateFilter,$DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {                    
                        Write-Verbose "[$($CA.Name)] Processing pending request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Pending' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
            Else
            {
                do 
                {
                    $Paging = Get-PendingRequest -CertificationAuthority $CertificateAuthority -Filter $DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {                   
                        Write-Verbose "[$($CA.Name)] Processing pending request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Pending' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
        }
        catch
        {
            $errorMsg = $_.Exception.Message
            $outcome = "Failed to query pending requests from CA. Exception: $errorMsg"
            Write-Log $outcome -LogPath $logPath
            Write-EventLog -Message $outcome -LogName Application -Source $eventSource -EntryType Error -EventId 9000
            Exit
        }
    }

    #======================================================================
    # Get failed certs
    #======================================================================
    if ($Failed)
    {
        Write-Log "[$($CA.Name)] Searching for failed certificate requests" -LogPath $logPath

        try 
        {
            $DateFilter = "Request.SubmittedWhen -ge $NotBeforeDate"
            $RequestIDFilter = "RequestID -ge $NotBeforeRequestID"
            $PageIndex = 1

            If ($CertificateTemplateOID)
            {
                do
                {
                    $TemplateFilter = "CertificateTemplate -eq $CertificateTemplateOID"
                    $Paging = Get-AdcsDatabaseRow -CertificationAuthority $CertificateAuthority -Table Failed -Filter $TemplateFilter,$DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {                    
                        Write-Verbose "[$($CA.Name)] Processing failed request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Failed' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
            Else
            {
                do
                {
                    $TemplateFilter = "CertificateTemplate -eq $CertificateTemplateOID"
                    $Paging = Get-AdcsDatabaseRow -CertificationAuthority $CertificateAuthority -Table Failed -Filter $DateFilter,$RequestIDFilter -PageSize $PageSize -Page $PageIndex -ErrorAction Stop | ForEach-Object {                                     
                        Write-Verbose "[$($CA.Name)] Processing failed request ID: $($_.RequestID)"
                        $_ | Select-Object *,
                            @{n='TemplateFriendlyName';e={ if($_.CertificateTemplateOid.Value -match "[a-z]"){ $_.CertificateTemplate} elseif ($_.CertificateTemplateOid.FriendlyName -eq $null){$_.CertificateTemplateOid.Value} else { $_.CertificateTemplateOid.FriendlyName }}},
                            @{n='Status';e={ 'Failed' }}
                    }

                    # increment page
                    $PageIndex++

                    # output current results
                    $Paging
                }
                while ($Paging -ne $null)
            }
        }
        catch
        {
            $errorMsg = $_.Exception.Message
            $outcome = "Failed to query failed requests from CA. Exception: $errorMsg"
            Write-Log $outcome -LogPath $logPath
            Write-EventLog -Message $outcome -LogName Application -Source $eventSource -EntryType Error -EventId 9000
            Exit
        }
    }
}
