#=================================================================================#
# PowerShell script sample for coolOrange powerJobs                               #
# Creates a PDF file and uploads it to to Sharepoint                              #
#                                                                                 #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                          #
#                                                                                 #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER       #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES     #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.      #
#=================================================================================#

# Install-Module -Name SharePointPnPPowerShellOnline -RequiredVersion 3.2.1810.0

# JobEntityType = FILE

#region Settings
# To include the Revision of the main file in the PDF name set Yes, otherwise No
$pdfFileNameWithRevision = $true

# The character used to separate file name and Revision label in the PDF name such as hyphen (-) or underscore (_)
$pdfFileNameRevisionSeparator = "_"

# To include the file extension of the main file in the PDF name set Yes, otherwise No
$pdfFileNameWithExtension = $true
# To add the PDF to Vault set Yes, to keep it out set No
$addPDFToVault = $true

# Specify a Vault folder in which the PDF should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$pdfVaultFolder = ""

# Specify a network share into which the PDF should be copied (e.g. \\SERVERNAME\Share\Public\PDFs\)
$pdfNetworkFolder = ""
# Please enter username to access your Sharepoint site
$Username = ""
# Please enter password to access your Sharepoint site
$Password = ConvertTo-SecureString -String "" -AsPlainText -Force
# Please enter the link of your Sharepoint site
$Site = ""
# Please enter in which folder you want to safe the files
$Folder = ""
#endregion


$pdfFileName = [System.IO.Path]::GetFileNameWithoutExtension($file._Name)
if ($pdfFileNameWithRevision) {
    $pdfFileName += $pdfFileNameRevisionSeparator + $file._Revision
}
if ($pdfFileNameWithExtension) {
    $pdfFileName += "." + $file._Extension
}
$pdfFileName += ".pdf"
if ([string]::IsNullOrWhiteSpace($pdfVaultFolder)) {
    $pdfVaultFolder = $file._FolderPath
}

try { 
    Connect-PnPOnline -Url $Site -Credential $PSCredentials 
    if (-not (Get-PnPContext)) { 
        Write-Host "Error connecting to SharePoint Online, unable to establish context" -foregroundcolor black -backgroundcolor Red 
        return 
    } 
} catch { 
    Write-Host "Error connecting to SharePoint Online: $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return 
}
if ( @("idw", "dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if (-not $addPDFToVault -and -not $pdfNetworkFolder) {
    throw("No output for the PDF is defined in ps1 file!")
}
if ($pdfNetworkFolder -and -not (Test-Path $pdfNetworkFolder)) {
    throw("The network folder '$pdfNetworkFolder' does not exist! Correct pdfNetworkFolder in ps1 file!")
}

[System.Management.Automation.PSCredential]$PSCredentials  = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($Username, $Password)



# Connect to Sharepoint
$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $file._ReleasedRevision

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."




$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | Select-Object -First 1
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen } 

if($openResult) {
    if($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    } else {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
    }                  
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile 

    if($exportResult) {
        $PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
    }
    
    Add-PnPFile -Path $localPDFfileLocation -Folder $Folder
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localPDFfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job '$($job.Name)'"