#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file with a Watermark on the bottom right corner              # 
# and add it to Autodesk Vault as Design Vizualization                        #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

#Please install the Add-Watermark cmdlet before using this script!!!
#https://support.coolorange.com/support/solutions/articles/22000216500

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
$HorizontalAlignment = "Buttom"
# To change the position vertical
$VerticalAlignment = "Right" 
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


Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."

if( @("idw") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | Select-Object -First 1
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen } 
$size=$openResult.document.instance.activesheet.size

$SizeMapping = @{
    "9986" = 50.0 #Custom Format
    "9987" = 30.0 #Format A
    "9988" = 40.0 #Format B
    "9989" = 50.0 #Format C
    "9990" = 60.0 #Format D
    "9991" = 70.0 #Format E
    "9992" = 80.0 #Format F
    "9993" = 80.0 #Format A0
    "9994" = 70.0 #Format A1
    "9995" = 60.0 #Format A2
    "9996" = 50.0 #Format A3
    "9997" = 40.0 #Format A4
    "9998" = 30.0 #Format 9 in x 12 in
    "9999" = 40.0 #Format 12 in x 18
    "10000" = 50.0 #Format 18 in x 24
    "10001" = 60.0 #Format 24 in x 36
    "10002" = 70.0 #Format 36 in x 48
    "10003" = 65.0 #Format 30 in x 42
}

foreach($i in $SizeMapping.Keys) {
    if($i -eq $size) {
        $FontSize = $SizeMapping[$i]
    }
}

if($openResult) {
    if($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    } else {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
    }                  
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile

    try {
        $text = $file._State
        Add-WaterMark -Path $localPDFfileLocation -WaterMark $text -Angle $Angle -HorizontalAlignment $HorizontalAlignment -VerticalAlignment $VerticalAlignment -Color $Color -Opacity $Opacity -FontSize $FontSize -OffSetX $OffsetX -OffSetY $OffsetY
    } catch [System.Exception] {
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
