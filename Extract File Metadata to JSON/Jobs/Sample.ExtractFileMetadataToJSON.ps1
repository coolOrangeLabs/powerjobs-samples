#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file and add it to Autodesk Vault as Design Vizualization     #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$workingDirectory = "C:\Temp\$($file._Name).json"

if(!(Test-Path "$workingDirectory")){
	New-Item -Path "$workingDirectory" -ItemType Directory | Out-Null
}

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."

$properties = @()
foreach ($prop in $file.PSObject.Properties){
	$properties += [ordered]@{
		"Name"=$prop.Name
		"Value"=$prop.Value
	}
}
$properties = ConvertTo-Json $properties

if(!($JSONfile)) {
	Set-Content -Path "$workingDirectory\$($file._Name).json" -Value $properties
} else {
	$JSONfile = New-Item -Path "$workingDirectory\$($file._Name).json" -ItemType File
	Set-Content -Path "$workingDirectory\$($file._Name).json" -Value $properties
}

Write-Host "Completed job '$($job.Name)'"