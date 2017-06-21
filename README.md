# Azure Automation Script that Tags Azure Resource Groups #

## Why
Automatically tag new Resource Group with an alias an expiry date, so that clean-up and responsibility become easier.

## What does it do
The Azure Automation Script does the following... 
* Find Resource Groups **without** the **alias** tag
* **Ignoring** certain standard RGs by a REGEX pattern supplied
* Search through *Azure Activity Logs* by using `Get-AzureRmLog` if there's any activities on these groups
* Tag the Resource Group with the **alias** if a user was **found in the logs**
* Tag the Resource Group with a **deleteAfter** date - 1 month in the future
* Send out an e-mail to people about the newly tagged RGs (so people know and can correct if necessary)

**Note:** The script can only search 14 days back of log, to find out who the owner is. Earlier activities cannot be considered.

*The script is not actually deleting anything ever. Just tagging with a deleteAfter date.*

## Deployment
* Make sure you have the HTML template and header graphics deployed to a public URL (e.g. Blob Storage with a Shared Access Signature)
* Deploy the script as an Azure Automation Script through the Azure Portal
* Set-up a scheduled job (https://docs.microsoft.com/en-us/azure/automation/automation-schedules)
* **Add the required parameters in all four sections**:

### 1. Azure Automation Runbook Inputs
| Parameter | Default   | Description                                               |
| --------- | -------   | -----------                                               |
| WhatIf    | False     | Goes into a dry-run mode of not actually tagging anything |
| To        | Empty     | Receipients of the mail, semicolon seperated and with <> brackets. e.g. "\<asd@asd.com\>;\<foo@foo.com\>" |
| DayCount  | 1         | How many days to look back in the logs (between 1-14) |

### 2. Azure Automation Variables
| Parameter | Description                                               |
| --------- | -----------                                               |
| RG_NamesIgnore    | Ignore these RG names (a regex) E.g. *(Default-\|AzureFunctions\|Api-Default-).\** |
| SubscriptionId | The subscription ID GUID to monitor |
| TemplateUrl    | The URL of the HTML Template used for the mail |
| TemplateHeaderGraphicUrl | The URL of the Header Graphic used for the mail (review script as this is optional) |

### 3. Azure Automation Credentials
| Name | Description     
| -----| ---------------|
| Office365 | Username / Password for the SMTP Server

### 4. Azure Automation Connection
| Name | Description     
| -----| ---------------|
| AzureRunAsConnection | Created by default
