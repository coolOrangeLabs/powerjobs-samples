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

# Required in the powerJobs Settings Dialog to determine the entity type for lifecycle state change triggers
# JobEntityType = FILE

#region Settings
$hidePDF = $false
$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath + "/" + (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $file._ReleasedRevision
$Color = "Orange"
$FontSize = 100
$HorizontalAlignment = "Center"
$VerticalAlignment = "Middle"
$Opacity = 100
$OffsetX = -2
$OffsetY = 15
$Angle = 170
#endregion Settings

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."

if ( @("idw", "dwg") -notcontains $file._Extension ) {
	Write-Host "Files with extension: '$($file._Extension)' are not supported"
	return
}

# Preloading the modules is required as PDFsharp 6.1.1 attempts to load the .NET stanard version of System.Runtime.CompilerServices.Unsafe.dll, wich causes errors in .NET Framework. v4 directory contains .NET Framework version of System.Runtime.CompilerServices.Unsafe.dll
Get-ChildItem -LiteralPath "$env:POWERJOBS_MODULESDIR\PDFWatermark" -Filter '*.dll' -Recurse | ForEach-Object {
	$null = [System.Reflection.Assembly]::LoadFile($_.FullName)
}
Import-Module ("$env:POWERJOBS_MODULESDIR\PDFWatermark\coolOrange.Pdf.WaterMark.dll")

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | Select-Object -First 1
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen }

if ($openResult) {
	if ($openResult.Application.Name -like 'Inventor*') {
		$configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
	}
	else {
		$configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg"
	}
	$exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile

	try {
		$text = $file._State
		Add-WaterMark -Path $localPDFfileLocation -WaterMark $text -FontSize $FontSize -Angle $Angle -HorizontalAlignment $HorizontalAlignment -VerticalAlignment $VerticalAlignment -Color $Color -Opacity $Opacity -OffSetX $OffsetX -OffSetY $OffsetY
	}
	catch [System.Exception] {
		throw($error[0])
	}


	if ($exportResult) {
		$PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization -Hidden $hidePDF
		$file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
	}
	$closeResult = Close-Document
}

if (-not $openResult) {
	throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if (-not $exportResult) {
	throw("Failed to export document $($file.LocalPath) to $localPDFfileLocation! Reason: $($exportResult.Error.Message)")
}
if (-not $closeResult) {
	throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job '$($job.Name)'"
