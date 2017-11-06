#Requires -Modules ActiveDirectory
###############################################################################
#                                                                             #
#            Get Active Directory User group membership quickly               #
#  (create CSV; membership column has list of groups separated by NewLines)   #
#                                                                             #
#                 NOTE:  probably doesn't scale very well :)                  #
#                                                                             #
###############################################################################
#                                                                             #
#  date:   2017-11-05                                                         #
#  note:   made with love by github.com/milesgratz                            #
#                                                                             #
###############################################################################

# Define list of AD users
$List = Get-Content "C:\temp\AD Users.txt"
$Output = "C:\temp\Results.csv"

# Create array for results and loop through list
$Results = @()
$List | ForEach-Object { 

    Try 
    { 
        # use 'Get-ADUser -Properties MemberOf' for faster lookup
        # convert DistinguishedName to Name
        $Membership = (Get-ADUser $_ -Properties MemberOf -ErrorAction Stop).MemberOf 
        $Membership = $Membership | ForEach-Object { (($_ -split "CN\=")[1] -split ",")[0] } | Out-String
    } 
    Catch 
    { 
        # add error message to Membership
        $Membership = "FAILURE: $($_.Exception.Message)"
    } 
    
    # Create custom object for results + add to array
    $Object = New-Object PSCustomObject 
    $Object | Add-Member -MemberType NoteProperty -Name "UserID" -Value $_
    $Object | Add-Member -MemberType NoteProperty -Name "MemberOf" -Value $Membership 
    $Results += $Object
}

# Output results to file
$Results | Export-Csv $Output -NoTypeInformation
