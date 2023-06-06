#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file with a Watermark sign across the drawing                 #
# and add it to Autodesk Vault as Design Vizualization                        #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

#Please install the Add-Watermark cmdlet before using this script!!!
#https://support.coolorange.com/a/solutions/articles/22000216500

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

# To enable faster opening of released Inventor drawings without downloading and opening their model files set Yes, otherwise No
$openReleasedDrawingsFast = $true
# To choose the color 
$Color = "Orange"
# To change the font size
$FontSize = 100
# To change the position horizontal
$HorizontalAlignment = "Center"
# To change the position vertical
$VerticalAlignment = "Middle" 
# To change the Opacity
$Opacity = 100
# To change the offset X
$OffsetX = -2
# To change the offset Y
$OffsetY = 15
# To change the angle
$Angle = 315

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
Write-Host "Starting job 'Create PDF as visualization attachment' for file '$($file._Name)' ..."

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
$fastOpen = $openReleasedDrawingsFast -and $file._ReleasedRevision


$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)

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

    try
    {
        $text = $file._State
        Add-WaterMark -Path $localPDFfileLocation -WaterMark $text -FontSize $FontSize -Angle $Angle -HorizontalAlignment $HorizontalAlignment -VerticalAlignment $VerticalAlignment -Color $Color -Opacity $Opacity -OffSetX $OffsetX -OffSetY $OffsetY
    }
        catch [System.Exception]
    {
        throw($error[0])
    }


    if($exportResult) {
        $PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
    }
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