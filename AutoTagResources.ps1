Param(
    [Parameter()]
    [boolean]$WhatIf = $False,
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter()]
    [ValidateRange(1, 14)] 
    [int32]$DayCount = 1
)

$days = $DayCount
if ($DayCount -gt 0) {
    $days = $DayCount * -1
}

$connectionName = "AzureRunAsConnection"
# the subscription ID of the Azure subscription 
$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"

# the template URL of the HTML Template used for the mail
$TemplateUrl = Get-AutomationVariable -Name "TemplateUrl"
$TemplateHeaderGraphicUrl = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl"

# ignore some resource groups (a REGEX - e.g. "(Default-|AzureFunctions|Api-Default-).*")
$RGNamesIgnoreRegex = Get-AutomationVariable -Name "RG_NamesIgnore"

# Credentials for sending the mail - name should be Office365
# https://docs.microsoft.com/en-us/azure/automation/automation-credentials
$mailCreds = Get-AutomationPSCredential -Name 'Office365'

# The mail server
$mailServer = "smtp.office365.com";

# Single Domain that users are in
$userdomain = "@microsoft.com";

# deletion date (just a warning in the mail and another TAG, no real delete here)
# 1 month in the future
$deleteDate = (Get-Date).AddMonths(1).ToString("MM\/dd\/yy")

try {
    # Get the connection "AzureRunAsConnection"
    $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         
    
    Write-Verbose "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null
    
    Set-AzureRmContext -SubscriptionId $SubscriptionId | Out-Null
}
catch {
    if (!$servicePrincipalConnection) {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    }
    else {
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

if ($WhatIf) {
    Write-Warning "Running in WhatIf mode - no changes will be made."
}

$allRGs = (Get-AzureRmResourceGroup).ResourceGroupName

Write-Verbose "Found $($allRGs.Length) total RGs"

$aliasedRGs = (Get-AzureRmResourceGroup -Tag @{ alias = $null }).ResourceGroupName

Write-Verbose "Found $($aliasedRGs.Length) aliased RGs"

$notAliasedRGs = $allRGs | ? {-not ($aliasedRGs -contains $_)}

Write-Verbose "Found $($notAliasedRGs.Length) un-tagged RGs"

$result = New-Object System.Collections.ArrayList

foreach ($rg in $notAliasedRGs) {
    if ($rg -match $RGNamesIgnoreRegex) {
        Write-Verbose "Ignoring Resource Group $rg"
        continue
    }

    $p = 100 / ($notAliasedRGs.Length - 1 ) * $notAliasedRGs.IndexOf($rg)
    Write-Progress -Activity "Searching Resource Group Logs for last $days days..." -PercentComplete $p `
        -CurrentOperation "$p% complete" `
        -Status "Resource Group $rg"

    $callers = Get-AzureRmLog -ResourceGroup $rg -DetailedOutput `
        -StartTime (Get-Date).AddDays($days) `
        -EndTime (Get-Date)`
        | Where-Object Caller -like "*@*" `
        | Where-Object { $_.Caller -and ($_.Caller -ne "System") } `
        | Where-Object { $_.OperationName.Value -ne "Microsoft.Storage/storageAccounts/listKeys/action" }`
        | Where-Object { $_.Properties.Content -and ($_.Properties.Content["requestbody"] -ne "{""tags"":{}}" ) } `
        | Sort-Object -Property Caller -Unique `
        | Select-Object Caller
    
    if ($callers) {
        $alias = $callers[0].Caller -replace $userdomain, ""

        if ($alias -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) {
            Write-Verbose "Ignoring guid alias $alias"
        }
        else {			
            Write-Verbose "Tagging Resource Group $rg for alias $alias"
            if (-not $WhatIf) {
                Set-AzureRmResourceGroup -Name $rg -Tag @{ alias = $alias; deleteAfter = $deleteDate} | Out-Null
            }
            $result.Add((New-Object PSObject -Property @{Name = $rg; Alias = $alias})) | Out-Null           
        }
    }
    else {
        Write-Verbose "No activity found for Resource Group $rg"
    }
}

Write-Progress -Activity "Searching Resource Group Logs..." -Completed -Status "Done"

if ($result.Count -gt 0) {
    $rgString = ($result | ForEach-Object { "<tr><td>$($_.Name)</td><td>$($_.Alias)</td></tr>" })

    $toAffected = ($result | ForEach-Object { "<$($_.Alias)$($userdomain)>" }) -join ";"

    $template = Invoke-WebRequest -Uri $TemplateUrl -UseBasicParsing

    $body = $template -replace "_TABLE_", $rgString -replace "_DATE_", $deleteDate

    $subject = "$($result.Count) new resource groups automatically tagged";

    $tocomb = "$To;$toAffected"

    # in WHATIF mode only send to certain users
    if ($WhatIf) {
        $tocomb = "$To"
    }

    $toArray = $tocomb.Split(";")

    Write-Verbose "Sending Mail to $tocomb"

    Invoke-WebRequest -UseBasicParsing $TemplateHeaderGraphicUrl -OutFile C:\template.png

    # This is specifically for having a tagging.png header graphic embedded to the mail
    # of course you can remove the line -InlineAttachments
    .\Send-MailMessageEx.ps1 `
        -Body $body `
        -Subject $subject `
        -Credential $mailCreds `
        -SmtpServer $mailServer `
        -Port 587 `
        -BodyAsHtml `
        -UseSSL `
        -InlineAttachments @{ "tagging.png" = "C:\template.png" } `
        -From $mailCreds.UserName `
        -To $toArray `
        -Priority "Low"
}
else {
    Write-Verbose "No Email sent - 0 Resource Groups tagged"
}

$result