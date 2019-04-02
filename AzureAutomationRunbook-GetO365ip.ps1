<# 
    .SYNOPSIS 
        Use Microsoft Graph API to check for changes in the Office 365 Ip Addresses.

    .DESCRIPTION 
        Calls to the below endoint return the version of IP address list
        https://endpoints.office.com/version?clientrequestid=<client request guid>
        
        Calls to the below endpoint return the actul list of IP address list
        https://endpoints.office.com/endpoints/worldwide?clientrequestid=<client request guid>

        Check version of the list. If current list version is greater than the stored version, perform the below work:
    .REQUIREMENTS
        O365 service account with a mailbox.
        O365 service account stored credentials in Azure Automation account.
    .AUTHOR 
        steveflowers@fastmail.com
    .VERSION 
        20190402
        
#> 

#region Variables
# Required static variables
$resourceGroupName = "automation_test"
$AutomationAccountName = "steve-test"
$listVersionURL = "https://endpoints.office.com/version"
$listDataURL = "https://endpoints.office.com/endpoints/worldwide"
$adminSmtpAddress = "steve.flowers@o-i.com"
$azureAutomaionCredentialName = "aa-ga"
$smtpServer = "smtp.office365.com"


# regex match IPv4 addresses ignoring IPv6
$regexIPv4 = "^([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9]|[1-2][0-9]|3[0-2]))?$"


$clientGuid = New-Guid
[string]$reqId = "?clientrequestid=$clientGuid"
$listVersionURL = "$listVersionURL" + "$reqId"
$listDataURL = "$listDataURL" + "$reqId"
#endregion

#region Get Azure Automation Credential
try{
    $azureAutomationCredential = Get-AutomationPSCredential -Name $azureAutomaionCredentialName
}
catch{
    $_
    Exit
}
#endregion

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
        <#
            .DETAILS
                Use Microsoft Graph API to get Office 365 IP addresses list version
            .STATUS
                Working
            .OUTPUT
                Int64
        #>
        try {
            $version = Invoke-RestMethod $URL
        }
        catch {
            Throw "Error retreiving IP list version from REST endpoint."

        }
        $version = $version | Where-Object instance -eq "Worldwide"
        return [int]$version.Latest
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
    Function notifyAdminOfChange ($adminSmtpAddress, $report)
    {
        $htmlhead = "<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <p>Office 365 IPs have been changed!</p>"

        $htmltail = "</body></html>"

        $html = $report | Out-String

        $body = $htmlhead + $html + $htmltail

        try{
            Send-MailMessage `
                -Credential $azureAutomationCredential `
                -To $adminSmtpAddress `
                -Body $body `
                -UseSsl `
                -SmtpServer $smtpServer
        }
        catch{
            $_
            Exit
        }
    }

#endregion


#region Get stored variables
try {
    $storedVariables = getStoredVariables
}
catch {
    $_
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
    $storedListData = ($storedVariables | Where-Object Name -eq "storedList").Value  | ConvertFrom-Json
}
catch {
    Write-Error $_
    Exit
}
#endregion

<#
    PSeudo code
    
    If stored list is empty, seed the data and stamp the version.
    If not, check the version:
        if 0, set the version and seed the list
        if > 0, compare stored version to new version
            if stored version -eq new version, do nothing
            if stored version -lt new version
                Get new ip list
                compare to stored list
                    Item in new list but not stored? That is an update
                    Item in stored list but not the new? That is a delete
                Set azure automation variable with new ip list
                Set azure automation variable version number
                Done
#>

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
        Set-AzureRmAutomationVariable `
                    -AutomationAccountName $AutomationAccountName `
                    -ResourceGroupName $resourceGroupName `
                    -Name "storedVersion" `
                    -Value $ipVersion `
                    -Encrypted $false
    }
    catch{
        $_
        Exit
    }
}
else {
    # Stored list is not empty
    # Get new list
    # Compare
    # Item in new list but not stored? That is an update
    # Item in stored list but not the new? That is a delete

    If ($storedListVersion -eq 0){
        try{
            $ipList = getOfficeIpData $listDataURL $regexIPv4
            $newIpVersion = getOfficeIpVersion $listVersionURL
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
            Set-AzureRmAutomationVariable `
                        -AutomationAccountName $AutomationAccountName `
                        -ResourceGroupName $resourceGroupName `
                        -Name "storedVersion" `
                        -Value $newIpVersion `
                        -Encrypted $false
        }
        catch{
            $_
            Exit
        }

    }
    elseif ($storedListVersion -eq $newIpVersion){
        #Nothing to do
        Write-Output "List has not changed."
        Exit
    }
    elseif ($storedListVersion -lt $newIpVersion){
        #List has updated
        try{
            $ipList = getOfficeIpData $listDataURL $regexIPv4
            $newIpVersion = getOfficeIpVersion $listVersionURL
        }
        catch{
            $_
            Exit
        }

        # Compare stored list to new list of IPs
        try{
            $compareResults = Compare-Object -ReferenceObject $storedListData -DifferenceObject $ipList
        }
        catch{
            $_
            Exit
        }
        
        # report is a pscustomobject that will be used in the email message body
        $report = @()

        foreach ($i in $compareResults){
            If ($i.SideIndicator -eq "=>"){
                $report += [PSCustomObject]@{
                    IP = $i.InputObject
                    Type = "Update"
                }
            }
            if ($i.SideIndicator -eq "<="){
                $report += [PSCustomObject]@{
                    IP = $i.InputObject
                    Type = "Delete"
                }
            }
        }


        # Store new data
        try{
            $ipListJson = $ipList | ConvertTo-Json
            Set-AzureRmAutomationVariable `
                        -AutomationAccountName $AutomationAccountName `
                        -ResourceGroupName $resourceGroupName `
                        -Name "storedList" `
                        -Value $ipListJson `
                        -Encrypted $false
            Set-AzureRmAutomationVariable `
                        -AutomationAccountName $AutomationAccountName `
                        -ResourceGroupName $resourceGroupName `
                        -Name "storedVersion" `
                        -Value $newIpVersion `
                        -Encrypted $false
        }
        catch{
            $_
            Exit
        }

        notifyAdminOfChange -adminSmtpAddress $adminSmtpAddress -report $report

    }
    else {
        Write-Error "Error enumerating list version. Exiting"
        Exit
    }


}
#endregion