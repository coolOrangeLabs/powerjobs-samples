#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a SAT file and adds it to Autodesk Vault as Design Substitute       #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$workingDirectory = "C:\Temp\$($file._Name)"
$localSTLfileLocation = "$workingDirectory\$($file._Name).STL"
$vaultSTLfileLocation = $file._EntityPath + "/" + (Split-Path -Leaf $localSTLfileLocation)

Write-Host "Starting job 'Create STL as attachment' for file '$($file._Name)' ..."

if ( @("ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory
$file = $downloadedFiles | Select-Object -First 1

$openResult = Open-Document -LocalFile $file.LocalPath #-Application InventorServer
if ($openResult) {
    try {
        $InvApp = $openResult.Application.Instance
        $STLAddin = $InvApp.ApplicationAddIns | Where-Object { $_.ClassIdString -eq "{533E9A98-FC3B-11D4-8E7E-0010B541CD80}" }
        $Context = $InvApp.TransientObjects.CreateTranslationContext()
        $Options = $InvApp.TransientObjects.CreateNameValueMap()

        if($openResult.Application.Name -eq "Inventor") {
            Write-Host "Using Inventor..."
            $SourceObject = $InvApp.ActiveDocument
        } elseif ($openResult.Application.Name -eq "InventorServer") {
            Write-Host "Using Inventor Server..."
            $SourceObject = $InvApp.Documents[1]
        } else {
            throw "$($openResult.Application.Name) not supported"
        }

        if ($STLAddin.HasSaveCopyAsOptions($SourceObject, $Context, $Options)) {
            $Options.Value("Resulution") = 1 #2=High, 1=Medium, 0=Low
            #$Options.Value("ExportUnits") = 4
            #$Options.Value("AllowMoveMeshNode") = $false
            #$Options.Value("SurfaceDeviation") = 60
            #$Options.Value("NormalDeviation") = 14
            #$Options.Value("MaxEdgeLength") = 100
            #$Options.Value("AspectRatio") = 40
            #$Options.Value("ExportFileStructure") = 0
            $Options.Value("OutputFileType") = 0 #0=binary, 1=ASCII
            #$Options.Value("ExportColor") = $true
        }

        $oData = $InvApp.TransientObjects.CreateDataMedium()
        $Context.Type = 13059       #kFileBrowseIOMechanism
        $oData.MediumType = 56577   #kFileNameMedium
        $oData.FileName = $localSTLfileLocation
        $STLAddin.SaveCopyAs($SourceObject, $Context, $Options, $oData)
        $exportResult = $true

        $STLfile = Add-VaultFile -From $localSTLfileLocation -To $vaultSTLfileLocation -FileClassification DesignDocument -Hidden $false
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($STLfile._FullPath)
    }
    catch {
        $exportResult = $false
        $exportResult | Add-Member -NotePropertyName Error -NotePropertyValue $_.Exception
    }
    
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory      

if (-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if (-not $exportResult) {
    throw("Export error. Failed to export STL document! Reason: $($exportResult.Error.Message)")
}
if (-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message)")
}
Write-Host "Completed job 'Create STL as attachment'"