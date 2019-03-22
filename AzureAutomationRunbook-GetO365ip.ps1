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

# Required static variables
$resourceGroupName = "automation_test"
$AutomationAccountName = "steve-test"
$listVersionURL = "https://endpoints.office.com/version"
$listDataURL = "https://endpoints.office.com/endpoints/worldwide"

# regex match IPv4 addresses ignoring IPv6
$regexIPv4 = "^([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9]|[1-2][0-9]|3[0-2]))?$"


# Initial seed of the list - comment out when done

# Gather stored variables
<#
$listDataEndpoint = Get-AzureRmAutomationVariable `
                                -ResourceGroupName $resourceGroupName `
                                –AutomationAccountName $AutomationAccountName `
                                -Name "listDataEndpoint"
#>
$clientGuid = New-Guid
[string]$reqId = "?clientrequestid=$clientGuid"
$listVersionURL = "$listVersionURL" + "$reqId"
$listDataURL = "$listDataURL" + "$reqId"




# Get the latest version
try
{
    $currentListVersion = Invoke-RestMethod $listVersionURL
    $currentListVersion = $currentListVersion | Where-Object instance -eq "Worldwide"
}
catch
{
    Throw "Error calling $listVersionURL."
}


If ($currentListVersion -gt $storedListVersion)
{
    # Get the latest list of IPs
    try
    {
        $currentListIPs = Invoke-RestMethod $listDataURL
        $currentListIPs = $currentListIPs.ips 
        $currentListIPs = $currentListIPs -match $regexIPv4
    }
    catch
    {
        Throw "Error calling $listVersionURL."
    }
}






