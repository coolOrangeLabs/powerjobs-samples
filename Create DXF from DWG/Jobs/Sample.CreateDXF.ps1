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

# To include the Revision of the main file in the dxf name set Yes, otherwise No
$dxfFileNameWithRevision = $true

# The character used to separate file name and Revision label in the dxf name such as hyphen (-) or underscore (_)
$dxfFileNameRevisionSeparator = "_"

# To include the file extension of the main file in the dxf name set Yes, otherwise No
$dxfFileNameWithExtension = $true

# To add the dxf to Vault set Yes, to keep it out set No
$addDXFToVault = $true

# To attach the dxf to the main file set Yes, otherwise No
$attachDCFToVaultFile = $true

# Specify a Vault folder in which the dxf should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$dxfVaultFolder = ""

# Specify a network share into which the PDF should be copied (e.g. \\SERVERNAME\Share\Public\PDFs\)
$dxfNetworkFolder = ""

# To enable faster opening of released Inventor drawings without downloading and opening their model files set Yes, otherwise No
$openReleasedDrawingsFast = $true

#endregion


$dxfFileName = [System.IO.Path]::GetFileNameWithoutExtension($file._Name)
if ($dxfFileNameWithRevision) {
    $dxfFileName += $dxfFileNameRevisionSeparator + $file._Revision
}
if ($dxfFileNameWithExtension) {
    $dxfFileName += "." + $file._Extension
}
$dxfFileName += ".dxf"

if ([string]::IsNullOrWhiteSpace($dxfVaultFolder)) {
    $dxfVaultFolder = $file._FolderPath
}

Write-Host "Starting job 'Create dxf as visualization attachment' for file '$($file._Name)' ..."


if( @("dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if (-not $adddxfToVault -and -not $dxfNetworkFolder) {
    throw("No output for the dxf is defined in ps1 file!")
}
if ($dxfNetworkFolder -and -not (Test-Path $pdfNetworkFolder)) {
    throw("The network folder '$dxfNetworkFolder' does not exist! Correct dxfNetworkFolder in ps1 file!")
}
if(!(Test-Path "$workingDirectory")){
	New-Item -Path "$workingDirectory" -ItemType Directory | Out-Null
}



$Title = $($file.Name).Substring(0,$($file.Name).LastIndexOf("."))
$localPNGfileLocation = "$workingDirectory\$($Title).dxf"
$vaultPNGfileLocation = $file._EntityPath +"/"+ (split-path -Leaf $localPNGfileLocation)
$fastOpen = $file._Extension  -eq "dwg" -and $file._ReleasedRevision

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."



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
$fastOpen = $openReleasedDrawingsFast -and $file._ReleasedRevision
$file = (Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen)[0]
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen }
if (-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if (-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localPDFfileLocation! Reason: $($exportResult.Error.Message)")
}
if (-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
if ($ErrorCopyPDFToNetworkFolder) {
    throw("Failed to copy PDF file to network folder '$staNetworkFolder'! Reason: $($ErrorCopyPDFToNetworkFolder)")
}

Write-Host "Completed job 'Create dxf as visualization attachment'"















































