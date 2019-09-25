#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a DWG2000 file and add it to Autodesk Vault as                      #
# Design Vizualization                                                        #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$hideDWG = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$localDWGfileLocation = "$workingDirectory\$($file._Name).dwg"
$vaultDWGfileLocation = $file._EntityPath +"/"+ (split-path -Leaf $localDWGfileLocation)
$fastOpen = $file._Extension -eq "idw" -and $file._ReleasedRevision

Write-Host "Starting job 'Create DWG2000 as attachment' for file '$($file._Name)' ..."

if( @("idw","dwg","iam","ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | Select-Object -First 1
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen }

if($openResult) {
    if(@("idw","dwg") -contains $file._Extension) {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_2D.ini" 
    } elseif ( $openResult.Document.Instance.ComponentDefinition.Type -eq [Inventor.ObjectTypeEnum]::kSheetMetalComponentDefinitionObject) {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_SheetMetal.ini" 
    } else { 
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWG_3D.ini" 
    }

    $exportResult = Export-Document -Format 'DWG' -To $localDWGfileLocation -Options $configFile -OnExport {
        param($export)

        $options = @{
            "DwgVersion" = 23
        }
        $export.Options = $options
    }

    if($exportResult) {
        $localDWGfiles = Get-ChildItem -Path (split-path -path $localDWGfileLocation) | Where-Object { $_.Name -ne $file._Name -and $_.Name -match '^'+[System.IO.Path]::GetFileNameWithoutExtension($localDWGfileLocation)+'.*(.dwg|.zip)$' }
        $vaultFolder = (Split-Path $vaultDWGfileLocation).Replace('\','/')
        $DWGfiles = @()
        foreach($localDWGfile in $localDWGfiles)  {
            $DWGfile = Add-VaultFile -From $localDWGfile.FullName -To ($vaultFolder+"/"+$localDWGfile.Name) -FileClassification DesignVisualization -Hidden $hideDWG
            $DWGfiles += $DWGfile._FullPath
        }
        $file = Update-VaultFile -File $file._FullPath -AddAttachments $DWGfiles
    }
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localDWGfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create DWG2000 as attachment'"