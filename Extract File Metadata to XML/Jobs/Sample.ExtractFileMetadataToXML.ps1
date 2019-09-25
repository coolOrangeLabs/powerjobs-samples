# ============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a XML file with the Metadata and uploads it to the UNC path         #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
# ============================================================================#

$workingDirectory = "C:\Temp\$($file._Name).xml"
[xml]$document = New-Object System.Xml.XmlDocument

if (!(Test-Path $workingDirectory)) {
    New-Item -Path $workingDirectory -ItemType Directory
}

Write-Host "Starting job '$($job.Name)' for file '$($file._Name)' ..."

if( @("idw","dwg") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
$decl = $document.CreateXmlDeclaration("1.0", "UTF-8", "yes")
$root = $document.CreateElement("Properties")

$document.AppendChild($root) | Out-Null
$document.InsertBefore($decl, $document.DocumentElement) | Out-Null



foreach ($prop in $file.PSObject.Properties) {
    $node = $document.CreateNode("element", "Property", "")
    $node.SetAttribute("Name", $($prop.Name))
    $node.InnerText = $prop.Value
    $root.AppendChild($node)
}

$document.Save("$workingDirectory\$($file.Name).xml") | Out-Null

Write-Host "Completed job '$($job.Name)'"