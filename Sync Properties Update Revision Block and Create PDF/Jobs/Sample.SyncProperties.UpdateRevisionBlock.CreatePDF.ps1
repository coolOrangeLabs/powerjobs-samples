# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#
#using Vault API to get the file instead of using powerVault's $file since it runs into property sync job not finding any equivalence problems when '$file.Name' or '$file._Name' is used
$fileForSyncJob = $vault.DocumentService.GetFileById($file.Id)

Write-Host "Starting job 'Sync Properties, Update Revision Block and create pdf' for file '$($fileForSyncJob.Name)' ..."

#update revision block job is added 'sometimes' after the sync job finishes. This can be switched off by removing the last two settings under the section <syncPropertiesPostJobExtensions> inside the JobProcessor.exe.config file.

$JobHandlerSyncPropertiesAssembly = [System.Reflection.Assembly]::Load("Connectivity.Explorer.JobHandlerSyncProperties")
$syncPropertiesJobHandler = $JobHandlerSyncPropertiesAssembly.CreateInstance("Connectivity.Explorer.JobHandler.SyncProperties.ExtHandlerSyncPropertiesJobHandler",$true, [System.Reflection.BindingFlags]::CreateInstance, $null, $null, $null, $null)

$syncPropertiesJob = New-Object Connectivity.Services.Job.SyncPropertiesJob("vault",[long]$fileForSyncJob.Id,$fileForSyncJob.Name,$false) # params - vaultname, fileversionid, filename, queueCreateDwfJobOnCompletion - there's overloaded constructor for working with collections

$ConnectivityJobProcessorDelegateAssembly = [System.Reflection.Assembly]::Load("Connectivity.JobProcessor.Delegate")
$context = $ConnectivityJobProcessorDelegateAssembly.CreateInstance("Connectivity.JobHandlers.Services.Objects.ServiceJobProcessorServices",$true, [System.Reflection.BindingFlags]::CreateInstance, $null, $null, $null, $null)

$context.GetType().GetProperty("Connection").SetValue($context, $vaultConnection, $null) #need to set the Connection property on the context or else it runs into error

$jobOutcome = $syncPropertiesJobHandler.Execute($context,$syncPropertiesJob) #call execute to start running the job

if ($jobOutcome -eq "Failure")
{
    throw "Failed job 'Sync Properties'" #Failed because of issue that occured in the job
}
Write-Host "Sync Properties completed"

Write-Host "Start Updating Revision Block"

$latestfile = $vault.DocumentService.GetLatestFileByMasterId($fileForSyncJob.MasterId)
$JobHandlerURBAssembly = [System.Reflection.Assembly]::Load("Connectivity.Explorer.JobHandlerUpdateRevisionBlock")
$uRBJobHandler = $JobHandlerURBAssembly.CreateInstance("Connectivity.Explorer.JobHandlerUpdateRevisionBlock.UpdateRevisionBlockJobHandler",$true, [System.Reflection.BindingFlags]::CreateInstance, $null, $null, $null, $null)
$uRBJob = New-Object Connectivity.Services.Job.UpdateRevisionBlockJob("vault",$latestfile.Id,$false,$false,$latestfile.Name)

$jobOutcome = $uRBJobHandler.Execute($context,$uRBJob) #call execute to start running the job

if ($jobOutcome -eq "Failure")
{
    throw "Failed job 'Update Revision Block'" #Failed because of issue that occured in the job
}

Write-Host "Finished Updating Revision Block"

$hidePDF = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $file._ReleasedRevision

Write-Host "'Create PDF as attachment' for file '$($file._Name)' ..."

if( @("idw","dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$file = $downloadedFiles | select -First 1
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen } 

if($openResult) {
    if($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    } else {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
    }                  
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile
    if($exportResult) {       
        $PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization -Hidden $hidePDF
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
    }
    $closeResult = Close-Document
}

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
Write-Host "Completed job 'Sync Properties, Update Revision Block and create PDF'"
