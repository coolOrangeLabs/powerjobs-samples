
$ErrorActionPreference = "stop"
# JobEntityType = ITEM
#region Settings

# Specify where the XML should be published (e.g. \\SERVERNAME\Share\Public\BOM push)
$xmlOutputFolder = "c:\temp\BOM push"
# Defines the XML export structure and mapping
$xmlTemplateFullPath = 'C:\ProgramData\coolOrange\powerJobs\Jobs\Sample.ItemBomPushToXmlTemplate.xml'

#endregion

New-Item -Path $xmlOutputFolder -ItemType Directory -Force

if (-not (Test-Path $xmlTemplateFullPath)) {                                
    throw "XML tamplate not found. Please check the path $($xmlTemplateFullPath)"
}
[XML]$xmlTemplate = Get-Content $xmlTemplateFullPath


function ConvertValue($property,$value) {
    if($null -ne $xmlTemplate.Root.ConvertionMapping.$property) {
        $valueMapping = Select-Xml -Xml $xmlTemplate.Root -XPath "//ConversionMapping/$property/ValueMapping[@Vault='$value']"
        if($null -ne $valueMapping) { return $valueMapping.Node.ERP }
        else {
            Write-Host "WARNING: Value mapping for '$value' on property '$property' could not be found in the value conversion table"
            return $value
        }
    }
    return $value    
}

Write-Host "Starting exporting BOM to XML"

$xmlName = "$($item._Number).xml"
$xmlOutputFullPath = "$($xmlOutputFolder)\$($xmlName)"
$bomheader = $item
[array]$itemBom = Get-VaultItemBOM -Number $item._Number

$xmlBomTemplate = $xmlTemplate.Root.BOMHeader

foreach ($headerProperty in $xmlBomTemplate.Property) {
    $propertyNameVault = $headerProperty.PropertyNameVault
    if ($headerProperty.'#Text' -ne "Value") { continue }
    $vaultValue = [string]$bomheader.$propertyNameVault
    $headerProperty.'#Text' = ConvertValue -value $vaultValue -property $propertyNameVault
}

if ($itemBom.Count -gt 1) {
    for ($i = 1; $i -lt $itembom.Count; $i++) {
        $clonedRow = $xmlBomTemplate.BOMRows.LastChild.CloneNode($true)
        $xmlBomTemplate.BOMRows.AppendChild($clonedRow)
    }
}
else 
{  
    $xmlBomTemplate.RemoveChild($xmlBomTemplate.BOMRows)
}

$counter = 0
foreach ($xmlRow in $xmlBomTemplate.BOMRows.BOMRow) {
    foreach ($xmlBomRowAttribute in $xmlRow.Attributes) {
        if($xmlBomRowAttribute.Value -ne ""){
            $xmlBomRowAttribute.Value = [string]$itembom[$counter].$($xmlBomRowAttribute.Value)
        }
    }
    foreach ($xmlRowProperty in $xmlRow.Property) {
        $propertyNameVault = $xmlRowProperty.PropertyNameVault
        if ($xmlRowProperty.'#Text' -ne "Value") { continue }
        $vaultValue = [string]$itembom[$counter].$propertyNameVault
        $xmlRowProperty.'#Text' = ConvertValue -value $vaultValue -property $propertyNameVault
    }
    $counter += 1
}

$newXmlBom = [xml]$xmlBomTemplate.OuterXml
$newXmlBom.Save($xmlOutputFullPath)
if (-not (Test-Path $xmlOutputFullPath)) {                                
    throw "Failed to save XML to path$($xmlOutputFullPath)"
}

Clean-Up -folder $workingDirectory
Write-Host "Export XML to $($xmlOutputFullPath) completed!"





