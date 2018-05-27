Param(
    [Parameter()]
    [boolean]$WhatIf = $False,
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter(HelpMessage = "Amount of days an RG is overdue expiry, for it to be included in the warning mail")]
    [ValidateRange(1, 180)] 
    [int32]$PastMaxExpiryDays = 1,
    [Parameter(HelpMessage = "Maximum amount of days an RG is set to expire in the future, after which it is included in the warning mail")]
    [ValidateRange(180,365)] 
    [int32]$FutureMaxExpiryDays = 180
)

$pastDays = $PastMaxExpiryDays * -1

$connectionName = "AzureRunAsConnection"
$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"

# the template URLs of the HTML Template used for the mail
$TemplateUrl = Get-AutomationVariable -Name "TemplateUrl-Cleanup"
$TemplateUrl_TooFar = Get-AutomationVariable -Name "TemplateUrl-CleanupTooFar"
$TemplateHeaderGraphicUrl = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl-Cleanup"

# ignore some resource groups (a REGEX - e.g. "(Default-|AzureFunctions|Api-Default-).*")
$RGNamesIgnoreRegex = Get-AutomationVariable -Name "RG_NamesIgnore"

# Credentials for sending the mail - name should be Office365
# https://docs.microsoft.com/en-us/azure/automation/automation-credentials
$mailCreds = Get-AutomationPSCredential -Name 'Office365'

# The mail server
$mailServer = "smtp.office365.com";

# Single Domain that users are in
$userdomain = "@microsoft.com";

try {
    # Get the connection "AzureRunAsConnection "
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

$allRGs = (Get-AzureRmResourceGroup)  | Select-Object ResourceGroupName, Tags

Write-Verbose "Found $($allRGs.Length) total RGs"

$deleteTagged = ($allRGs | ? { $_.Tags.deleteAfter })

Write-Verbose "Found $($deleteTagged.Length) tagged RGs"

$notDeleteTagged = ($allRGs | ? { -not $_.Tags.deleteAfter })

Write-Verbose "Found $($notDeleteTagged.Length) un-tagged RGs"

$deleteTaggedCasted = $deleteTagged | Select-Object @{name = "DeleteAfter"; expression = {[datetime]$_.Tags.deleteAfter}}, `
    @{name = "Alias"; expression = {$_.Tags.alias}}, `
    @{name = "ResourceGroupName"; expression = {$_.ResourceGroupName}}, `
    @{name = "ResourceCount"; expression = { 0 }}, `
    @{name = "Resources"; expression = { @() }}
					
$expired = $deleteTaggedCasted | Where-Object {$_.DeleteAfter -lt (Get-Date).AddDays($pastDays)} `
    | Sort-Object -Property Alias `
    | Where-Object -Property ResourceGroupName -NotMatch $RGNamesIgnoreRegex

foreach ($item in $expired) {
    Write-Verbose "Fetching count for group $($item.ResourceGroupName)"
    $item.Resources = Get-AzureRmResource -ResourceGroupName $item.ResourceGroupName
    $item.ResourceCount = $item.Resources.Count
}

if ($expired.Count -gt 0) {
    # add an entry to the HTML table
    $rgString = ($expired | ForEach-Object { "<tr><td>$($_.ResourceGroupName)</td><td>$($_.Alias)</td><td>$($_.DeleteAfter)</td><td>$($_.ResourceCount)</td></tr>" })

    # add to the list of affected mails
    $toAffected = ($expired | ForEach-Object { "<$($_.Alias)$($userdomain)>" }) -join ";"

    # download HTML template from the web
    $template = Invoke-WebRequest -Uri $TemplateUrl -UseBasicParsing
    # download the header graphics
    Invoke-WebRequest -UseBasicParsing $TemplateHeaderGraphicUrl -OutFile C:\template.png

    # replace parameters in the template
    $body = $template -replace "_TABLE_", $rgString -replace "_DATE_", $deleteDate

    $subject = "$($expired.Count) resource groups expired";

    if ($WhatIf) {
        Write-Warning "WHATIF set, only sending to $To"
        $tocomb = "$To"
    }
    else {
        $tocomb = "$To;$toAffected"
    }

    $toArray = $tocomb.Split(";")

    Write-Verbose "Sending Mail to $tocomb"

    # Send the e-mail using the external script
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
    Write-Warning "No Email sent - 0 Resource Groups expired"
}

# filter for RGs that have an expiry dates further out than X days
$tooFarOutExpiry = $deleteTaggedCasted | Where-Object {$_.DeleteAfter -gt (Get-Date).AddDays($FutureMaxExpiryDays)} `
    | Sort-Object -Property DeleteAfter

foreach ($item in $tooFarOutExpiry) {
    Write-Verbose "Fetching count for group $($item.ResourceGroupName)"
    # fetch the list and count of resources
    $item.Resources = Get-AzureRmResource -ResourceGroupName $item.ResourceGroupName
    $item.ResourceCount = $item.Resources.Count
}

if ($tooFarOutExpiry.Count -gt 0) {
    # add an entry to the HTML table
    $rgString = ($tooFarOutExpiry | ForEach-Object { "<tr><td>$($_.ResourceGroupName)</td><td>$($_.Alias)</td><td>$($_.DeleteAfter)</td><td>$($_.ResourceCount)</td></tr>" })
    
    # add to the list of affected mails
    $toAffected = ($tooFarOutExpiry | ForEach-Object { "<$($_.Alias)$($userdomain)>" }) -join ";"
    
    # download HTML template from the web
    $template = Invoke-WebRequest -Uri $TemplateUrl_TooFar -UseBasicParsing      
    # download the header graphics
    Invoke-WebRequest -UseBasicParsing $TemplateHeaderGraphicUrl -OutFile C:\template.png
    
    # replace parameters in the template
    $body = $template -replace "_TABLE_", $rgString -replace "_DATE_", $deleteDate
    
    $subject = "$($tooFarOutExpiry.Count) Resource Groups have Expiry Date > 6 Months";
        
    $tocomb = "$To"
    
    $toArray = $tocomb.Split(";")
    
    Write-Verbose "Sending Mail about too far expiry to $tocomb"
    
    # Send the e-mail using the external script
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
    Write-Warning "No Email sent - 0 Resource Groups with too far expiry"
}

$expired