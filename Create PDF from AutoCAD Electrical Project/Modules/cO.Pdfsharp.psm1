<#
This module provides functionality to merge multiple pdf files. It was tested with PdfSharp-gdi.dll in version 1.5.4 and powershell 5 and is provided "as is"
#>
function MergePdf {
<#
.SYNOPSIS
Merges multiple PDF files into a single multisheet PDF
.PARAMETER Files
A collection of PDF files
.PARAMETER DestinationFile
Full path of the merged PDF
.PARAMETER PdfSharpPath
Full path to the required PdfSharp library
.PARAMETER Force
Tries to force destination file and directory creation and deletion of source files, even when they are read-only
.PARAMETER
RemoveSourceFiles
Deletes the source files after PDF is merged
.EXAMPLE
$files = Get-ChildItem "C:\temp\PDF\Source" -Filter "*.pdf"
MergePdf -Files $files  -DestinationFile "C:\TEMP\PDF\Destination\test.pdf" -PdfSharpPath 'C:\ProgramData\coolOrange\powerJobs\Modules\PdfSharp-gdi.dll' -Force -RemoveSourceFiles
#>
param(
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[ValidateScript({
    if( $_.Extension -ine ".pdf" ){ 
        throw "The file $($_.FullName) is not a pdf file."
    } 
    if(-not (Test-Path $_.FullName)) {
        throw "The file '$($_.FullName)' does not exist!"
    }
    $true
})]
[System.IO.FileInfo[]]$Files,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[System.IO.FileInfo]$DestinationFile,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
$PdfSharpPath,
[switch]$Force,
[switch]$RemoveSourceFiles
)
    Write-Host ">> $($MyInvocation.MyCommand.Name) >>"

    if((Test-Path $PdfSharpPath) -eq $false) {
        throw "Could not find pdfsharp assembly at $($PdfSharpPath)"
    }
    Add-Type -LiteralPath $PdfSharpPath

    if((Test-Path $DestinationFile.FullName) -and $DestinationFile.IsReadOnly -and -not $Force) {
        throw "Destination file '$($DestinationFile.FullName)' is read only"
    }

    [System.IO.DirectoryInfo]$DestinationDirectory = $DestinationFile | Split-Path -Parent
    if(-not (Test-Path $DestinationDirectory)) {
        try {
            $DestinationDirectory = New-Item -Path $DestinationDirectory.FullName -ItemType Directory -Force:$Force
        } catch {
            throw "Error in $($MyInvocation.MyCommand.Name). Could not create directory '$($Path)'. $Error[0]"
        }
    }

    $pdf = New-Object PdfSharp.Pdf.PdfDocument
    Write-Host "Creating new PDF"
    foreach ($file in $Files) {
        $inputDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($file.FullName, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
        for ($index = 0; $index -lt $inputDocument.PageCount; $index++) {
            $page = $inputDocument.Pages[$index]
            $null = $pdf.AddPage($page)
        }
    }

    Write-Host "Saving PDF"
    if((Test-Path $DestinationFile.FullName) -and $Force) { 
        Remove-Item $DestinationFile.FullName -Force 
    }
    $pdf.Save($DestinationFile.FullName)

    if($RemoveSourceFiles) {
        Write-Host "Removing source files"
        foreach($file in $files) {
            Remove-Item -Path $file.FullName -Force:$Force
        }
    }
}
