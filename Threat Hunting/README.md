# Introduction 
This script were created by Chaim Black. 

## Get-365InboxRules.ps1

Get-365InboxRules was created to audit user inbox rules in Office 365. 

Attackers typically create malicious inbox rules after compromising an account, often redirecting, forwarding, moving, or deleting mail. 
This script captures all user inbox rules, looks at several attributes which are often misused by attackers, and exports it to a xlsx (default) or CSV. It also outputs the raw results to json.

As a reminder, always check user inbox rules when investigating an email compromise. When running this script in larger companies, I often find either active email compromises, or remnants of prior compromises which was not fully cleaned up.


### Installation ##

Get-365InboxRules requires the following modules to be installed:
  - Connecting to Office 365:

ExchangeOnlineManagement https://www.powershellgallery.com/packages/ExchangeOnlineManagement

The script does not install this module and requires you to be connected to Exchange Online prior to connecting.

  - ImportExcel https://github.com/dfinke/ImportExcel

If not installed, the script will automatically install ImportExcel unless specifying the "ExportCSV" switch. 

Massive thank you to Doug Finke for creating ImportExcel!

### Usage ##

This script was created as a function and can either be pasted into PowerShell and then called, or loaded manually by referencing the file.

```
 . .\Get-365InboxRules.ps1
 Get-365InboxRules 
```

 Parameter | Possible Values 
--- | --- |
-NoLaunch | [Switch] (Optional) - Prevents the output folder from launching after completion of command. Default is set to open folder.
-Path | [String] (Optional) - Sets the path of the output. Default path is C:\PSOutput\Get-365InboxRules\
-ExportCSV | [Switch] (Optional) - By default, this script exports as an Excel file. Setting this switch exports as CSV instead.
-OutObject | [Switch] (Optional) -  Outputs all rules. Can be used if you want to pipe it to something else.
-OutRawObject | [Switch] (Optional) - Outputs all raw rules. Can be used if you want to pipe it to something else.
-Username | [Switch] (Optional) -By default, this script searches an entire Office 365 tenant. You can use this to specify a single user to search instead.

### Output
The script either outputs an XLSX file (default) or two CSV files, and a JSON file with the raw output.

In its default output of XLSX file, it generates a main worksheet with results, and if it detects known suspicious rules, it will put them in another worksheet named "Suspicious" 

If the "ExportCSV" switch is defined, it will instead output a main worksheet with the results, and if it detects known suspicious rules, it will put them in another CSV file named "Suspicious" 

It will also export the raw unfiltered data as a JSON file.

#### WARNING
I was hesitant to list any rules as suspicious since attackers are constantly changing their tactics, and I don't want you to rely on just this label and not look at all the rules.

Please use your brain when looking at the rule and think what an attacker can do. With the default output of XLSX, it easily allows you to filter between different columns such as looking at any rules which can delete, forward or move to strange folders.

Column explanation:
Column | Explanation 
--- | --- |
User | Username
Name | DisplayName
Enabled | Defines if rule is enabled
Delete | Defines if rule can delete mail
ApplyAll | Defines if rule can apply to all incoming mail
Date | Defines if rule can apply to emails before or after a date
Size | Defines if rule can be dependent to size of mail
Move | Defines if rule can move mail to another folder
MarkAsRead | Defines if rule can mark mail as read
FwdorRedir | Defines if rule can forward or redirect mail
FwdorRedirExt | Defines if rule can forward or redirect mail to an external domain
Rule | Defines rule
FwdorRedirTo | Defines where mail rule is forwarding to (if applicable)
FwdorRedirExtTo |Defines where mail rule is forwarding to if recipient is an external domain (if applicable)
UserObjectID | Defines User ObjectID

### Issues

Known Issues:
This script uses 'Get-InboxRule' which requires a mailbox to be specified. This script uses a foreach loop to capture the data for all mailboxes but can run into throttling issues if there are too many mailboxes.

### Legal Disclaimer

Use this script at your own risk. This script is offered “as-is” and does not come with any warranty, either expressed or implied. I take no reponsibility for any harm this can cause.

Copyright (c) 2021 Chaim Black
Please feel free to share this script.
