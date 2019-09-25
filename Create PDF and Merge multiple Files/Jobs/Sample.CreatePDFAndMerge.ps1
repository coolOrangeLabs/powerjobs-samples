#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file of multiple PDF files									  # 
# and add it to Autodesk Vault as Design Vizualization  					  #
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

Add-Type -Path "C:\ProgramData\coolOrange\powerJobs\Modules\PdfSharp.dll"
$files = Get-VaultFiles -Folder $vaultFolder
$files = $files | Where-Object { $_.'File Extension' -Match "^(idw|dwg)" }
$hideZIP = $false
$workingDirectory = "C:\Temp\Drawings"

Write-Host "Starting job '$($job.Name)' ..."

if(!(Test-Path "$workingDirectory\Export")){
    New-Item -Path "$workingDirectory\Export" -ItemType Directory | Out-Null
}

foreach ($file in $files){
	$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $files._ReleasedRevision
	$fastOpen = $false

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
$path="$workingDirectory\Export"
$outfile = "$workingDirectory\$([Guid]::NewGuid()).pdf"

$output = New-Object PdfSharp.Pdf.PdfDocument            

foreach($i in (Get-ChildItem -Path $path *.pdf)) {            
    $input = New-Object PdfSharp.Pdf.PdfDocument            
    $input = [PdfSharp.Pdf.IO.PdfReader]::Open($i.fullname, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)            
	for ($i = 0; $i -lt $input.PageCount; $i++){
		$output.AddPage($input.Pages[$i])
	}
}
$output.Save($outfile) 
Add-VaultFile -From $outfile -To $($files[0]._EntityPath + "/" + (Split-Path -Leaf $outfile)) -FileClassification DesignVisualization -Hidden $hideZIP

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