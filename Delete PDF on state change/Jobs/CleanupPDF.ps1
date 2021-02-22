#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file and add it to Autodesk Vault as Design Vizualization     #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#


#TODO: To copy the PDF to a network folder fill $networkFolder with the folder e.g. $networkFolder ="\\SERVER1\Share\Public\PDFs\"
$networkFolder = "<YOUR PATH>"
$localPdfFileName = "$($file._Name).pdf"

Write-Host "Starting job 'Cleanup PDF' for file '$($file._Name)' ..."

if (-not $networkFolder) {
    throw("ERROR: No output for the PDF is defined in ps1 file!")
}

if ( @("idw", "dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
$localPdfFullFileName  = [System.IO.Path]::Combine($networkFolder, $localPdfFileName)

Write-Host "Deleting file '$($localPdfFileName)'..."

if (Test-Path $localPdfFullFileName) {
    try {
        Remove-Item -Path $localPdfFullFileName
    }
    catch {
        throw "File cannot be removed (is it open?)"
    }
} else {
    Write-Host "File doesn't exist"
}

Write-Host "Completed job 'Publish PDF'"