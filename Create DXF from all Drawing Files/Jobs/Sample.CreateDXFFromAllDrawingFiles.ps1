#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates multiple DXF files of all Inventor-files and saves them             #
# in the Vault                                                                #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$files = Get-VaultFiles -Properties  @{"File Extension"="dwg"} 
$files += Get-VaultFiles -Properties    @{"File Extension"="idw"} 
$files = $files | Where-Object { $_.'File Extension' -Match "^(idw|dwg)" } #| Select-Object -First 10
 
$workingDirectory = "C:\Temp\DXF"
 
if(!(Test-Path "$workingDirectory")){
    New-Item -Path "$workingDirectory" -ItemType Directory
}
 
$vaultfolder = "$/Designs"
Write-Host "Starting job 'Create DXF of all Inventor files'$($files._Name)'  ..."
 
$accoreconsolepath = Resolve-Path -Path "C:\Program Files\Autodesk\*\accoreconsole.exe"
 
foreach ($file in $files) {
    $fastOpen = $file._Extension -eq "dwg" -or $file._Extension -eq "idw"-and $file._ReleasedRevision
    Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen | Out-Null
    $Title = [System.IO.Path]::GetFileNameWithoutExtension($file._FullPath)
    $localInventorFileLocation = "$workingDirectory\$Title.dxf"
    $vaultDWGFileLocation = $vaultfolder+"/DXF/"+ (split-path -Leaf "$Title.dxf")

    if($file._Extension -contains "idw"){
        $openResult = Open-Document -LocalFile "$workingDirectory\$($file._Name)" -Options @{ FastOpen = $fastOpen }

        if($openResult) {
            if(@("idw","dwg") -contains $file._Extension) {
                $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_2D.ini" 
            } elseif ( $openResult.Document.Instance.ComponentDefinition.Type -eq [Inventor.ObjectTypeEnum]::kSheetMetalComponentDefinitionObject) {
                $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_SheetMetal.ini" 
            } else { 
                $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_3D.ini" 
            }

            $exportResult = Export-Document -Format 'DWG' -To "$workingDirectory\$Title.dwg" -Options $configFile
            if(-not $exportResult) {
                throw("Failed to export document $($file.LocalPath) to $localDWGfileLocation! Reason: $($exportResult.Error.Message)")
            }
        }

        $closeResult = Close-Document
        if(-not $closeResult) {
            throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
        }
    } else {
        # Do not delete Break-space
        $text = "DXFOUT ""$localInventorFileLocation"" 16
        
        "
        $text > "$workingDirectory\script.scr"

        & $accoreconsolepath[0] /i "$workingDirectory\$Title.dwg" /s "$workingDirectory\script.scr" #| Out-Null
    }
 
    
    Add-VaultFile -From $localInventorFileLocation -To $vaultDWGFileLocation -FileClassification DesignVisualization -Hidden $false | Out-Null
    Write-Host "Processing file '$($file._Name)'..."
}
 
Write-Host "Completed job 'Create DXF for all Inventor files'"
Clean-Up -folder $workingDirectory