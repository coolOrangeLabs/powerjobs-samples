#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PNG of the Thumbnail 										      #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$workingDirectory = "C:\Temp\Thumbnails"

Write-Host "Starting job 'Create Thumbnail' from file '$($file._Name)' ..."

if(!(Test-Path $workingDirectory)){
	New-Item -Path $workingDirectory -ItemType Directory | Out-Null
}

$thumb = $file.Thumbnail

$imageBytes = $thumb.Image
$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$ms.Write($imageBytes, 0, $imageBytes.Length);
$image = [System.Drawing.Image]::FromStream($ms, $true)
$image.Save("$workingDirectory\$($file._Name).png")

Write-Host "Completed job 'Create Thumbnail'"