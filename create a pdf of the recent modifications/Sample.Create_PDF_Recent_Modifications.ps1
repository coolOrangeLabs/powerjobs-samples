Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\cO.Pdfsharp.psm1"

# JobEntityType = FILE

#region Settings
# smp server for what email system is used
$Server = 'smtp.office365.com'
# To include the file extension of the main file in the PDF name set Yes, otherwise No
$youremail = "julian.piazzi@coolorange.com"
# the email addres that get's the email
$useremail = "moritz.pruenster@coolorange.com"
# password to send a mail over your email
$password = "Qwer12-/"
# To add the PDF to Vault set Yes, to keep it out set No
$addPDFToVault = $true
# vault directiory
$pdfVaultFolder = 'C:\Users\MoritzPrünster\Documents\Vault\Designs\test'

# To attach the PDF to the main file set Yes, otherwise No
$email_head = "ein Furz"

# Specify a Vault folder in which the PDF should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$email_body = "arsch"

# To create a own name for the pdf
$pdfFileName = "filelist"
#endregion

#$workingDirectory = "C:\Temp"
# only to run the file without vault started
#if ( -not $iamrunninginjobprocessor ) {
   # $workingDirectory = "C:\Temp"
 #   Import-Module powerVault
#}

#----------------------------create table------------------------------
# Get the current date and calculate the date 7 days ago
$currentDate = Get-Date
$startDate = $currentDate.AddDays(-7)


# Get all files in the directory that have been modified within the last 7 days
$changedFiles = Get-ChildItem -Path $pdfVaultFolder -Recurse | Where-Object {
    $_.LastWriteTime -ge $startDate -and $_.LastWriteTime -le $currentDate
}
Write-Host $_.LastWriteTime
# Create a report file
$reportPath = "$workingDirectory\updated.txt"

# Generate the report content
$reportContent = "File Name`tLast Modified"
$reportContent += "`n" + ("-" * 50)

foreach ($file in $changedFiles) {
    $reportContent += "`n" + $file.Name + "`t" + $file.LastWriteTime.ToString()
}

# Save the report to the file
New-Item -ItemType Directory -Path $workingDirectory -Force
$reportContent | Out-File -FilePath $reportPath

# Display a success message
Write-Host "Report generated successfully. File saved at: $reportPath"

#------------------------create a pdf of the png file-----------------------------------

$linesPerGroup = 40

$imageFiles = @()
$pdfFiles = @()
$lines = @()
# Iterate through every line of the file
foreach ($line in Get-Content -Path $reportPath) {
    $lines += $line
}

$i = 0

while ($lines.Length - ($i * $linesPerGroup) -gt 0) {
    $start = $i * $linesPerGroup
    $txtContent = ""
    $imageFilePath = "$workingDirectory\$i.png"
    
    $pdfFilePath = "$workingDirectory\$i.pdf"
    $imageFiles += $imageFilePath
    $pdfFiles += $pdfFilePath
    for ($j = 0; $j -lt 40 -or $lines.Length - $start - $j -gt 0; $j++) {
        $txtContent += $lines[$start + $j] + "`n"
    }
    # Create a new bitmap image with increased resolution
    $bitmapWidth = 1200
    $bitmapHeight = 1800
    $bitmapResolution = 400
    $bitmap = New-Object System.Drawing.Bitmap($bitmapWidth, $bitmapHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $bitmap.SetResolution($bitmapResolution, $bitmapResolution)

    # Create a graphics object from the bitmap
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    # Create a font and brush for drawing text
    $font = New-Object System.Drawing.Font("Arial", 7)
    $brush = [System.Drawing.Brushes]::Black

    # Draw the text onto the bitmap using graphics object, font, and brush
    $graphics.DrawString($txtContent, $font, $brush, 10, 10)

    # Save the bitmap as a PNG image file
    $bitmap.Save($imageFilePath, [System.Drawing.Imaging.ImageFormat]::Png)

    # Dispose of the objects
    $graphics.Dispose()
    $bitmap.Dispose()

    # Create a PrintDocument object
    $printDoc = New-Object System.Drawing.Printing.PrintDocument

    # Define the PrintPage event handler
    $printDoc.add_PrintPage({
            param($sender, $e)

            # Load the image
            $image = [System.Drawing.Image]::FromFile($imageFilePath)

            # Set the print area to fit the entire image
            $e.Graphics.DrawImage($image, $e.MarginBounds)

            # Specify that there are no more pages to print
            $e.HasMorePages = $false

            # Dispose of the image
            $image.Dispose()
        })

    # Set the printer name to "Microsoft Print to PDF"
    $printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF"

    # Print to a file by setting the PrintToFile property to true
    $printDoc.PrinterSettings.PrintToFile = $true

    # Set the output file name
    $printDoc.PrinterSettings.PrintFileName = $pdfFilePath

    # Print the document
    $printDoc.Print()

    $i ++
}

#------------------create one pdf out of many pdf's----------------------------
$pdfFile = Get-Item "$workingDirectory\*.pdf"
$pdfFileName=$pdfFileName+".pdf"
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\cO.Pdfsharp.psm1"
MergePdf -PdfSharpPath "C:\ProgramData\coolOrange\powerJobs\Modules\Pdf\PdfSharp.dll" -Files $pdfFile -DestinationFile "$workingDirectory\$pdfFileName"
$localPDFfileLocation="$workingDirectory\$pdfFileName"
Write-Host "pdf created"

#------------------add to the vault folder------------------------------------------
if ($addPDFToVault) {
    #$pdfVaultFolder = $pdfVaultFolder.TrimEnd('/')
    Write-Host "Add PDF '$pdfFileName' to Vault: $pdfVaultFolder"
    $PDFfile = Add-VaultFile -From $localPDFfileLocation -To "$pdfVaultFolder\$pdfFileName" -FileClassification DesignVisualization
}

#-----------------send email------------

$password = ConvertTo-SecureString $password -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential ($youremail, $password)
Send-MailMessage -Credential $cred -from $youremail -to $useremail -Subject $email_head -SmtpServer $Server -Port '587'-UseSsl -Attachments $localPDFfileLocation -Body $email_body
    

Write-Host "finish job"
