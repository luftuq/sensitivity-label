Add-Type -AssemblyName System.IO.Compression.FileSystem

# Define the folder path and file to work with
$folderPath = Get-Location
$filenameToRemove = 'docProps/custom.xml'


$xmlContentConfidential = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
...
"@

$xmlContentInternal = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
...
"@

$xmlContentPublic = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
...
"@

$xmlContentRestricted = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
...
"@
# Get all xlsx files in the folder
$xlsxFiles = Get-ChildItem -Path $folderPath -Filter *.xlsx

$fileToInsert = $xmlContentInternal

foreach ($file in $xlsxFiles) {
    # Create a new temporary zip file
    $tempZip = [System.IO.Path]::GetTempFileName()

    # Change the extension of the temporary file to .zip
    $tempZip = [System.IO.Path]::ChangeExtension($tempZip, 'zip')

    # Copy the original xlsx file to the temporary zip file
    Copy-Item -Path $file.FullName -Destination $tempZip -Force

    # Open the zip file
    $zip = [System.IO.Compression.ZipFile]::Open($tempZip, 'Update')

    # Remove the specified file from the zip if it exists
    $entry = $zip.GetEntry($filenameToRemove)
    if ($entry) {
        $entry.Delete()
    }

    # Add the new file to the zip
    $entry = $zip.CreateEntry('docProps/custom.xml')
    $writer = New-Object System.IO.StreamWriter($entry.Open())
    $writer.Write($fileToInsert)
    $writer.Flush()
    $writer.Close()

    # Close the zip file
    $zip.Dispose()

    # Move the temporary zip file back to the original location and rename it to the original file name
    Move-Item -Path $tempZip -Destination $file.FullName -Force
}