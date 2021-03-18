# powerjobs-samples

[![Windows](https://img.shields.io/badge/Platform-Windows-lightgray.svg)](https://www.microsoft.com/en-us/windows/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5-blue.svg)](https://microsoft.com/PowerShell/)
[![Vault](https://img.shields.io/badge/Autodesk%20Vault-2020-yellow.svg)](https://www.autodesk.com/products/vault/)

[![powerJobs](https://img.shields.io/badge/coolOrange%20powerJobs-20-orange.svg)](https://www.coolorange.com/en-eu/enhance.html#powerJobs)

## Disclaimer

THE SAMPLE CODE ON THIS REPOSITORY IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.

THE USAGE OF THIS SAMPLE IS AT YOUR OWN RISK AND **THERE IS NO SUPPORT** RELATED TO IT.

## Description

Sample jobs for coolOrange powerJobs

This repository contains various jobs that demonstrate the power and flexibility of *coolOrange powerJobs*. As a certified partner, you can use the jobs as samples or templates to speed up your project development/customization and to deliver high quality and flexible jobs to your customer with less effort and less time.

## Debugging

In order to obtain a file from Vault and with that running the script in a code editor such as 'Windows PowerShell ISE' or 'Visual Studio Code' instead of powerJobs, the following code can be added to the ps1 scripts.

```powershell
if (-not $IAmRunningInJobProcessor) {
    Import-Module powerJobs
    Open-VaultConnection -Server "localhost" -Vault "Vault" -User "Administrator" -Password ""
    $file = Get-VaultFile -Properties @{Name="Scissors.idw"}
}
```

This additional code logs in to Vault and uses the file 'Scissors.idw' if the script gets exectued by anthing other than powerJobs.

## At your own risk
The usage of these samples is at your own risk. There is no free support related to the samples. However, if you have questions to powerJobs, then visit http://www.coolorange.com/wiki or start a conversation in our support forum at http://support.coolorange.com/support/discussions

## Author
coolOrange s.r.l.  

![coolOrange](https://i.ibb.co/NmnmjDT/Logo-CO-Full-colore-RGB-short-Payoff.png)
