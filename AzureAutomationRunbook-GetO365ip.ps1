<# 
    .SYNOPSIS 
        Use Microsoft Graph API to check for changes in the Office 365 Ip Addresses.

    .DESCRIPTION 
        Calls to the below endoint return the version of IP address list
        https://endpoints.office.com/version?clientrequestid=<client request guid>
        
        Calls to the below endpoint return the actul list of IP address list
        https://endpoints.office.com/endpoints/worldwide?clientrequestid=<client request guid>

        Check version of the list. If current list version is greater than the stored version, perform the below work:
    .AUTHOR 
        steveflowers@fastmail.com
    .VERSION 
        20190402
        
#> 

# Required static variables
$resourceGroupName = "automation_test"
$AutomationAccountName = "steve-test"
$listVersionURL = "https://endpoints.office.com/version"
$listDataURL = "https://endpoints.office.com/endpoints/worldwide"


# regex match IPv4 addresses ignoring IPv6
$regexIPv4 = "^([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9]|[1-2][0-9]|3[0-2]))?$"


$clientGuid = New-Guid
[string]$reqId = "?clientrequestid=$clientGuid"
$listVersionURL = "$listVersionURL" + "$reqId"
$listDataURL = "$listDataURL" + "$reqId"

#region Functions
    Function getStoredVariables
    {
        try {
            $vars = Get-AzureRmAutomationVariable `
                                        -ResourceGroupName $resourceGroupName `
                                        -AutomationAccountName $AutomationAccountName `
                                        -ErrorAction Stop
        }
        catch {
            Throw "Error retreiving stored variables."
        }
        return $vars

    }
    Function getOfficeIpVersion ($URL)
    {
        try {
            $version = Invoke-RestMethod $URL
        }
        catch {
            Throw "Error retreiving IP list version from REST endpoint."

        }
        $version = $version | Where-Object instance -eq "Worldwide"
        return $version
    }
    Function getOfficeIpData ($URL, $regex)
    {
        <#
            .DETAILS
                Use Microsoft Graph API to get Office 365 IP addresses
            .STATUS
                Working
            .OUTPUT
                Object string array
        #>

        try {
            $data = Invoke-RestMethod $URL
        }
        catch {
            $_
            Throw "Error retreiving IP list from REST endpoint."
        }
        $data = $data.ips
        $data = $data -match $regexIPv4
        return $data
    }

#endregion


#region Get stored variables
try {
    $storedVariables = getStoredVariables
}
catch {
    Write-Error $_ + "-------"
    Exit
}
#endregion


#region Check for stored variables, if missing, create
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
#endregion


#region Initialize stored variables
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
#endregion

#region Store List Data
if ($storedListData -eq "") {
    # If empty, intiial seed of data
    # Get list using function
    # convert to json
    # store in azure automation variable
    try{

        $ipList = getOfficeIpData $listDataURL $regexIPv4

    }
    catch{
        $_
        Exit
    }

    try{

        $ipListJson = $ipList | ConvertTo-Json
        Set-AzureRmAutomationVariable `
                    -AutomationAccountName $AutomationAccountName `
                    -ResourceGroupName $resourceGroupName `
                    -Name "storedList" `
                    -Value $ipListJson `
                    -Encrypted $false
    }
    catch{
        $_
        Exit
    }

}
else {
    # Get old list
    # Get new list
    # Compare
    # Item in new list but not old? That is an update
    # Item in old list but not the new? That is a delete

}
#endregion




# Get the latest version
getOfficeIpVersion -URL $listVersionURL

If ($currentListVersion -gt $storedListVersion)
{
    # Get the latest list of IPs
    getOfficeIpData -URL $listDataURL -regex $regexIPv4
}






