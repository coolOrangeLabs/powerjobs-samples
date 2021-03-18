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

$hidePDF = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $file._ReleasedRevision
$Color = "Orange"
$HorizontalAlignment = "Right" 
$VerticalAlignment = "Bottom" 
$Opacity = 50
$OffsetX = -2
$OffsetY = 0
$Angle = 0

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
        $PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization -Hidden $hidePDF
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
