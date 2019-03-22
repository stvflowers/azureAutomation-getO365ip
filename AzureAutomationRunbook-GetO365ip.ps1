<# 
    .SYNOPSIS 
        Use Microsoft Graph API to check for changes in the Office 365 Ip Addresses.

    .DESCRIPTION 
        Calls to the below endoint return the version of IP address list
        https://endpoints.office.com/version?clientrequestid=<client request guid>
        
        Calls to the below endpoint return the actul list of IP address list
        https://endpoints.office.com/endpoints/worldwide?clientrequestid=<client request guid>

        Check version of the list. If current list version is greater than the stored version, perform the below work:




    .PARAMETER AzureCredentialName 
        
    .PARAMETER AzureSubscriptionName 
        
    .PARAMETER Simulate 
        
    .EXAMPLE 
        
    .INPUTS 
        
    .OUTPUTS 
        
#> 

# stored variables
[string]$listVersionEndpoint = ""
[string]$listDataEndpoint = ""
[int]$storedListVersion = ""
$storedList = @()

$clientGuid = New-Guid
[string]$reqId = "?clientrequestid=$clientGuid"
$listVersionURL = "$listVersionEndpoint" + "$reqId"
$listDataURL = "$listDataEndpoint" + "$reqId"
[int]$currentListVersion = ""
$currentList = @()



# Initial seed of the list - comment out when done

# Gather stored variables
$listVersionEndpoint = Get-AzureRmAutomationVariable `
                                -ResourceGroupName "automation_test" `
                                –AutomationAccountName "steve-test" `
                                -Name 'listVersionEndpoint'
$listDataEndpoint


# Get the latest version
$currentListVersion = Invoke-RestMethod $listVersionEndpoint
try
{
    $currentListVersion = $currentListVersion | Where-Object instance -eq "Worldwide"
}
catch
{
    Throw "Error calling $listVersionEndpoint."
}


If ($currentListVersion -gt $storedListVersion)
{

}






