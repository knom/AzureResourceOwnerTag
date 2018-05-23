Param(
    [Parameter()]
    [boolean]$WhatIf = $False,
    [Parameter(Mandatory=$true)]
    [string]$To,
    [Parameter()]
    [ValidateRange(14,180)] 
    [int32]$DayCount = 32
)

$days = $DayCount
if ($DayCount -gt 0)
{
    $days = $DayCount * -1
}

$connectionName = "AzureRunAsConnection"
$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"
$TemplateUrl = Get-AutomationVariable -Name "TemplateUrl-Cleanup"
$RGNamesIgnoreRegex = Get-AutomationVariable -Name "RG_NamesIgnore"
$TemplateHeaderGraphicUrl = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl-Cleanup"

try
{
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection=Get-AutomationConnection -Name $connectionName         
    
    Write-Verbose "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null
    
    Set-AzureRmContext -SubscriptionId $SubscriptionId | Out-Null
}
catch {
    if (!$servicePrincipalConnection)
    {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    } else{
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

$allRGs = (Get-AzureRmResourceGroup)  | Select-Object ResourceGroupName, Tags

Write-Warning "Found $($allRGs.Length) total RGs"

$deleteTagged = ($allRGs | ? { $_.Tags.deleteAfter })

Write-Warning "Found $($deleteTagged.Length) tagged RGs"

$notDeleteTagged = ($allRGs | ? { -not $_.Tags.deleteAfter })

Write-Warning "Found $($notDeleteTagged.Length) un-tagged RGs"

$deleteTaggedCasted = $deleteTagged | Select-Object @{name="DeleteAfter";expression={[datetime]$_.Tags.deleteAfter}}, `
					@{name="Alias";expression={$_.Tags.alias}}, `
					@{name="ResourceGroupName";expression={$_.ResourceGroupName}}
					
$expired = $deleteTaggedCasted | Where-Object {$_.DeleteAfter -lt (Get-Date).AddDays(-31)} | Sort-Object -Property Alias `
    | Where-Object -Property ResourceGroupName -NotMatch $RGNamesIgnoreRegex

if ($expired.Count -gt 0)
{
    #$rgString = ($result | ForEach-Object { "$($_.Name) ($($_.Alias))" }) -join "<br/>"
    $rgString = ($expired | ForEach-Object { "<tr><td>$($_.ResourceGroupName)</td><td>$($_.Alias)</td><td>$($_.DeleteAfter)</td></tr>" })

    $toAffected = ($expired | ForEach-Object { "<$($_.Alias)@microsoft.com>" }) -join ";"

    $template = Invoke-WebRequest -Uri $TemplateUrl -UseBasicParsing

    $body = $template -replace "_TABLE_",$rgString -replace "_DATE_",$deleteDate

    #$body = "<p>Hi,<br>The following resource groups have been automatically tagged:</p>$($rgString)<p>Please verify and otherwise correct the alias tag!</p>"
    $subject = "$($expired.Count) Resource Groups Expired";

    $tocomb = "$To;$toAffected"

    $toArray = $tocomb.Split(";")

    Write-Warning "Sending Mail to $tocomb"

    $mailCreds = Get-AutomationPSCredential -Name 'Office365'
    #Send-MailMessage -Body $body -BodyAsHtml -Credential $mailCreds `
    #    -From $mailCreds.UserName `
    #    -Port 587 -SmtpServer smtp.office365.com -Subject $subject -To $toArray -UseSSL

    Invoke-WebRequest -UseBasicParsing $TemplateHeaderGraphicUrl -OutFile C:\template.png

    .\SendMailMessageEx.ps1 `
		-Body $body `
		-Subject $subject `
		-Credential $mailCreds `
		-SmtpServer smtp.office365.com `
		-Port 587 `
		-BodyAsHtml `
		-UseSSL `
		-InlineAttachments @{ "tagging.png" = "C:\template.png" } `
		-From $mailCreds.UserName `
		-To $toArray `
		-Priority "Low"
}
else{
    Write-Warning "No Email sent - 0 Resource Groups expired"
}

$result