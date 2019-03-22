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

#region Functions
    Function getStoredVariables
    {
        try {
            $storedVariables = Get-AzureRmAutomationVariable `
                                        -ResourceGroupName $resourceGroupName `
                                        –AutomationAccountName $AutomationAccountName
        }
        catch {
            Throw "Error retreiving stored variables."
        }
        return $storedVariables

    }
    Function getOfficeIpVersion ($URL)
    {
        try {
            $version = Invoke-RestMethod $URL
        }
        catch {
            Throw Write-Error "Error retreiving IP list version from REST endpoint."

        }
        $version = $version | Where-Object instance -eq "Worldwide"
        return $version
    }
    Function getOfficeIpData ($URL, $regex)
    {
        try {
            $data = Invoke-RestMethod $URL
        }
        catch {
            Throw  Write-Error "Error retreiving IP list from REST endpoint."
        }
        $data = $data.ips
        $data = $data -match $regexIPv4
        return $data
    }

#endregion

# Get stored variables
try {
    $storedVariables = getStoredVariables
}
catch {
    Write-Error $_
    Exit
}


# Check for stored variables, if missing, create
If (-not($($storedVariables.Name) -like "*storedVersion*" ))
{
    try {
        New-AzureRmAutomationVariable `
                    -ResourceGroupName $resourceGroupName `
                    -AutomationAccountName $AutomationAccountName `
                    -Name "storedVersion" `
                    -Description "Stored version of Office 365 IP list. The version numbers can be found at https://endpoints.office.com/version." `
                    -Value 0 `
                    -Encrypted $false
        $storedVariables = getStoredVariables
    }
    catch {
        Write-Error $_
        Exit
    }

}
If (-not($($storedVariables.Name) -like "*storedList*" ))
{
    try {
        New-AzureRmAutomationVariable `
                    -ResourceGroupName $resourceGroupName `
                    -AutomationAccountName $AutomationAccountName `
                    -Name "storedList" `
                    -Description "Stored list of Office 365 IP list. The list can be found at https://endpoints.office.com/endpoints/worldwide." `
                    -Value "" `
                    -Encrypted $false
        $storedVariables = getStoredVariables
    }
    catch {
        Write-Error $_
        Exit
    }

}

try {
    $storedListVersion = ($storedVariables | Where-Object Name -eq "storedVersion").Value
}
catch {
    Write-Error $_
    Exit
}

try {
    $storedListData = ($storedVariables | Where-Object Name -eq "storedList").Value
}
catch {
    Write-Error $_
    Exit
}

if ($storedListData -eq "") {
    # Is empty, intiial seed

}
else {

    
}




# Get the latest version
getOfficeIpVersion -URL $listVersionURL

If ($currentListVersion -gt $storedListVersion)
{
    # Get the latest list of IPs
    getOfficeIpData -URL $listDataURL -regex $regexIPv4
}






