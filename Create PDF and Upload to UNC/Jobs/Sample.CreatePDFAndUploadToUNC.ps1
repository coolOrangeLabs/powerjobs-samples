#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file and uploads it to a UNC path                             #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#
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

# To enter the unc path
$uncPath = ""

# Specify a Vault folder in which the PDF should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$pdfVaultFolder = ""

# Specify a network share into which the PDF should be copied (e.g. \\SERVERNAME\Share\Public\PDFs\)
$pdfNetworkFolder = ""

# To enable faster opening of released Inventor drawings without downloading and opening their model files set Yes, otherwise No
$openReleasedDrawingsFast = $true
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

$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $openReleasedDrawingsFast -and $file._ReleasedRevision
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

    Copy-Item -Path $localPDFfileLocation -Destination $uncPath
    
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
