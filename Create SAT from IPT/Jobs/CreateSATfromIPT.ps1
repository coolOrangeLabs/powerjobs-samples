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
# Required in the powerJobs Settings Dialog to determine the entity type for lifecycle state change triggers
# JobEntityType = FILE

#region Settings
# To include the Revision of the main file in the SAT name set Yes, otherwise No
$satFileNameWithRevision = $true

# The character used to separate file name and Revision label in the SAT name such as hyphen (-) or underscore (_)
$satFileNameRevisionSeparator = "_"

# To include the file extension of the main file in the SAT name set Yes, otherwise No
$satFileNameWithExtension = $true

# To add the SAT to Vault set Yes, to keep it out set No
$addSATToVault = $true

# Specify a Vault folder in which the SAT should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$satVaultFolder = ""

# Specify a network share into which the PDF should be copied (e.g. \\SERVERNAME\Share\Public\PDFs\)
$satNetworkFolder = ""

# To enable faster opening of released Inventor drawings without downloading and opening their model files set Yes, otherwise No
$openReleasedDrawingsFast = $true

#endregion

$localSATfileLocation = "$workingDirectory\$($file._Name).SAT"

$satFileName = [System.IO.Path]::GetFileNameWithoutExtension($file._Name)
if ($satFileNameWithRevision) {
    $satFileName += $satFileNameRevisionSeparator + $file._Revision
}
if ($satFileNameWithExtension) {
    $satFileName += "." + $file._Extension
}
$satFileName += ".SAT"

if ([string]::IsNullOrWhiteSpace($satVaultFolder)) {
    $satVaultFolder = $file._FolderPath
}

Write-Host "Starting job 'Create SAT as visualization attachment' for file '$($file._Name)' ..."


if ( @("ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if (-not $addSATToVault -and -not $satNetworkFolder) {
    throw("No output for the SAT is defined in ps1 file!")
}
if ($satNetworkFolder -and -not (Test-Path $satNetworkFolder)) {
    throw("The network folder '$satNetworkFolder' does not exist! Correct satNetworkFolder in ps1 file!")
}
#


$vaultSATfileLocation = $file._EntityPath + "/" + (Split-Path -Leaf $localSATfileLocation)
Write-Host "Starting job 'Create SAT as attachment' for file '$($file._Name)' ..."
$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory
$file = $downloadedFiles | Select-Object -First 1


#
$fastOpen = $openReleasedDrawingsFast -and $file._ReleasedRevision
$file = (Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen)[0]
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen }

if ($openResult) {
    if ($openResult.Document.Instance.ComponentDefinition.Type -ne [Inventor.ObjectTypeEnum]::kSheetMetalComponentDefinitionObject) {
        Write-Host "Part file is not a sheet metal part!"
        $exportResult = $true
    }
    else {
            $componentDefinition = $openResult.Document.Instance.ComponentDefiniton
            if ($componentDefinition.HasFlatPattern)
            {
                $componentDefinition.Unfold()
                $componentDefinition.FlatPattern.ExitEdit()
            }

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
            Write-Host "Completed job 'Create SAT as attachment'"
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
if ($ErrorCopyPDFToNetworkFolder) {
    throw("Failed to copy PDF file to network folder '$staNetworkFolder'! Reason: $($ErrorCopyPDFToNetworkFolder)")
}

Write-Host "Completed job 'Create SAT as visualization attachment'"