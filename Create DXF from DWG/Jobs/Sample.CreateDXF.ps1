#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a DFX file and add it to Autodesk Vault as Design Vizualization     #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$workingDirectory = "C:\Temp\$($file._Name)"
$Title = $($file.Name).Substring(0,$($file.Name).LastIndexOf("."))
$localPNGfileLocation = "$workingDirectory\$($Title).dxf"
$vaultPNGfileLocation = $file._EntityPath +"/"+ (split-path -Leaf $localPNGfileLocation)
$fastOpen = $file._Extension  -eq "dwg" -and $file._ReleasedRevision

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."

if( @("dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if(!(Test-Path "$workingDirectory")){
	New-Item -Path "$workingDirectory" -ItemType Directory | Out-Null
}
# Not Delete Break-space
$text = "DXFOUT

"
$text > "$workingDirectory\script.scr"
$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | Select-Object -First 1
$accoreconsolepath = Resolve-Path -Path "C:\Program Files\Autodesk\*\accoreconsole.exe"
& $accoreconsolepath[0]  /i "$workingDirectory\$($file._Name)" /s "$workingDirectory\script.scr"
$DWGfile = Add-VaultFile -From $localPNGfileLocation -To $vaultPNGfileLocation -FileClassification DesignVisualization -Hidden $false
$file = Update-VaultFile -File $file._FullPath -AddAttachments @($DWGfile._FullPath)

Clean-Up -folder $workingDirectory
Write-Host "Completed job '$($job.Name)'"