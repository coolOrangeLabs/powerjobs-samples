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
$localSATfileLocation = "$workingDirectory\$($file._Name).SAT"
$vaultSATfileLocation = $file._EntityPath + "/" + (Split-Path -Leaf $localSATfileLocation)

Write-Host "Starting job 'Create SAT as attachment' for file '$($file._Name)' ..."

if ( @("ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory
$file = $downloadedFiles | Select-Object -First 1

$openResult = Open-Document -LocalFile $file.LocalPath
if ($openResult) {
    if ($openResult.Document.Instance.ComponentDefinition.Type -ne [Inventor.ObjectTypeEnum]::kSheetMetalComponentDefinitionObject) {
        Write-Host "Part file is not a sheet metal part!"
        $exportResult = $true
    }
    else {
        $openResult.Document.Instance.ComponentDefiniton.Unfold
         
        try {
            $InvApp = $openResult.Application.Instance
            $SATAddin = $InvApp.ApplicationAddIns | Where-Object { $_.ClassIdString -eq "{89162634-02B6-11D5-8E80-0010B541CD80}" }
            $SourceObject = $InvApp.ActiveDocument
            $Context = $InvApp.TransientObjects.CreateTranslationContext()
            $Options = $InvApp.TransientObjects.CreateNameValueMap()
            $oData = $InvApp.TransientObjects.CreateDataMedium()
            $Context.Type = 13059       #kFileBrowseIOMechanism
            $oData.MediumType = 56577   #kFileNameMedium
            $oData.FileName = $localSATfileLocation
            $SATAddin.SaveCopyAs($SourceObject, $Context, $Options, $oData)
            $exportResult = $true

            $SATfile = Add-VaultFile -From $localSATfileLocation -To $vaultSATfileLocation -FileClassification DesignSubstitute -Hidden $false
            $file = Update-VaultFile -File $file._FullPath -AddAttachments @($SATfile._FullPath)
        }
        catch {
            $exportResult = $false
        }
    }
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory      

if (-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if (-not $exportResult) {
    throw("Export error. Inventor failed to export SAT document to $($vaultSATfileLocation)")
}
if (-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create SAT as attachment'"