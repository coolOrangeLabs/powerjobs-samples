#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates multiple PDF files and compresses the files to a ZIP archive        #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

if (-not $folder) {
    $vaultFolder = "$/Designs/Inventor Sample Data/Models/Assemblies/Scissors" 
} else {
    $vaultFolder = $folder._FullPath
}

$files = Get-VaultFiles -Folder $vaultFolder 
$files = $files | Where-Object { $_.'File Extension' -Match "^(idw|dwg)" }
$hideZIP = $false
$workingDirectory = "C:\Temp\$([Guid]::NewGuid())"

Write-Host "Starting job '$($job.Name)' ..."

if(!(Test-Path "$workingDirectory\Export")){
    New-Item -Path "$workingDirectory\Export" -ItemType Directory | Out-Null
}

foreach ($file in $files){
    $fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $files._ReleasedRevision
    $downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory "$workingDirectory\Import" -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
    $file = $downloadedFiles | Select-Object -First 1
    $localPDFfileLocation = "$workingDirectory\Export\$($file._Name).pdf"
    
    $openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen } 
     if($openResult) {
        if($openResult.Application.Name -like 'Inventor*') {
               $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
        } else {
            $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
        }                  
        $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile
    
        $closeResult = Close-Document
    }
}
$ZIPfile = "$workingDirectory\$([Guid]::NewGuid()).zip"
Compress-Archive -Path "$workingDirectory\Export" -DestinationPath $ZIPfile  -Force

Add-VaultFile -From $ZIPfile -To $($files[0]._EntityPath + "/" + (Split-Path -Leaf $ZIPfile)) -FileClassification DesignVisualization -Hidden $hideZIP
    
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