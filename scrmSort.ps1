# scrmSort
# original author: Chris Finlay
# For sorting SCRM localised delivery packages 
# v0.6.4
# Changelog at end of file

##########################
## Function Definitions ##
##########################
function Copy-Languages {
    param ($itemName)

    Write-Host "#####################" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host -NoNewLine "Breakdown of " 
    Write-Host $itemName -ForegroundColor Yellow -BackgroundColor Black
    $partsOfName = $itemName -split("\.")
    $fileName = $partsOfName[0]
    $fileExtension = $partsOfName[1]
    $zoneNameAndID = $fileName -split("-")
    $zoneName = $zoneNameAndID[0]
    $secondParams = $zoneNameAndID[1] -split("_")
    $uniqueID = $secondParams[0]
    
    
    # Testing to see what parts we end up with after splitting secondParams up. uniqueID is the first part (0), after that we may have 0, 1 or 2 elements to the name (global English, global language or licence specific language)
    if ($secondParams.Count -gt 1) {
        if ($secondParams.Count -eq 2) {
            $language = $secondParams[1]
            if ($language -eq "GLOBAL") {
                #If the language solely states "GLOBAL", it's EN - and needs to be renamed to GLOBAL_EN
                Write-Host "GLOBAL" -ForegroundColor Blue -BackgroundColor Yellow
                $language="EN"
                $newFileName = $fileName+"_EN."+$fileExtension
                Rename-Item $file.fullname $newFileName
                # $file still contains the reference to the previous filename, so we need to make sure to try and copy the renamed file, not the old file (which no longer exists and would cause PS to throw an error)
                # Problem here is that the Copy-Item operation needs to use $file.fullPath... which we no longer have access to
                # So we split the Path of the old filename to get the directory, and then tack it on to the start of the new filename. Then we pass that entire string into the Copy-Item operation
                $fileDirectory = Split-Path $file.fullname
                $newFullPath = $fileDirectory+"\"+$newFileName
                Write-Host "This is a global language file."
                Write-Host -NoNewLine "The language is "
                Write-Host $language -ForegroundColor Black -BackgroundColor DarkCyan
                $destinationDirectory = $zoneName+"-"+$uniqueID+"-GLOBAL"
                if (-not (Test-Path $outputLocation\_finalOutput\$destinationDirectory)){
                    mkdir $outputLocation\_finalOutput\$destinationDirectory | out-null 
                }
                Copy-Item $newFullPath $outputLocation\_finalOutput\$destinationDirectory
            }
            else {
                Write-Host "This is a global language file."
                Write-Host -NoNewLine "The language is "
                Write-Host $language -ForegroundColor Black -BackgroundColor DarkCyan
                $destinationDirectory = $zoneName+"-"+$uniqueID
                if (-not (Test-Path $outputLocation\_finalOutput\$destinationDirectory)){
                    mkdir $outputLocation\_finalOutput\$destinationDirectory | out-null 
                }
                Copy-Item $file.fullname $outputLocation\_finalOutput\$destinationDirectory
            }
        }
        ElseIf ($secondParams.Count -eq 3) {
            Write-Host "This is a specific language file."
            $licence = $secondParams[1]
            $language = $secondParams[2]
            
            Write-Host -NoNewLine "The language is "
            Write-Host $language -ForegroundColor Black -BackgroundColor DarkCyan
            Write-Host -NoNewLine "The licence is "
            Write-Host $licence -ForegroundColor Black -BackgroundColor DarkCyan
            $licenceUpper = $licence.ToUpper()
            $destinationDirectory = $zoneName+"-"+$uniqueID+"-"+$licenceUpper
            if (-not (Test-Path $outputLocation\_finalOutput\$destinationDirectory)){
                mkdir $outputLocation\_finalOutput\$destinationDirectory | out-null 
            }
            Copy-Item $file.fullname $outputLocation\_finalOutput\$destinationDirectory
        }
        else {
            Write-Host "Something has gone wrong. The following file will be discarded." -ForegroundColor Red
            Write-Host "Please check the name and run the script again if necessary." -ForegroundColor Red
            Write-Host $file.fullname -ForegroundColor Red
            Pause
        }
    }
    else {
        Write-Host "This is a global English file."
        $destinationDirectory = $zoneName+"-"+$uniqueID
        $fileWithoutExtension = $file -split("\.")
        $newName = $fileWithoutExtension[0]+"_en.html"
        Rename-Item -Path $file.fullname -NewName $newName
        if (-not (Test-Path $outputLocation\_finalOutput\$destinationDirectory)){
            mkdir $outputLocation\_finalOutput\$destinationDirectory | out-null 
        }
        Copy-Item $outputLocation\_working\$newName $outputLocation\_finalOutput\$destinationDirectory
    }
    Write-Host -NoNewLine "File name is "
    Write-Host -NoNewLine $fileName -ForegroundColor Black -BackgroundColor Green
    Write-Host -NoNewLine ", extension "
    Write-Host $fileExtension -ForegroundColor Black -BackgroundColor Green
    Write-Host -NoNewLine "Zone name is "
    Write-Host -NoNewLine $zoneName -ForegroundColor Black -BackgroundColor DarkCyan
    Write-Host -NoNewLine ", unique ID is "
    Write-Host $uniqueID -ForegroundColor Black -BackgroundColor DarkCyan
} 

function Compress-Results {
# Compresses the final copied results into zip files
Write-Host "Compressing files for final delivery" -ForegroundColor White -BackgroundColor Green
    $mainDirectoryContents = @(Get-ChildItem $outputLocation\_finalOutput)
    foreach ($subDirectory in $mainDirectoryContents) {
        $shortFilesToZip = @(Get-ChildItem $outputLocation\_finalOutput\$subDirectory)
        foreach ($fileToZip in $shortFilesToZip) {
            $longFileToZip = $fileToZip.fullname
            $fileSplit = $fileToZip -split("\.")
            $fileToZipWithoutExtension = $fileSplit[0]
            $compress = @{
                Path = $longFileToZip
                DestinationPath = $outputLocation+"\_finalOutput\"+$subDirectory+"\"+$fileToZipWithoutExtension+".zip"
            }
            Compress-Archive @compress
            Remove-Item $fileToZip.fullname
        }
        $compress = @{
            Path = $subDirectory.fullname+"\*"
            DestinationPath = $outputLocation+"\_finalOutput\"+$subDirectory+".zip"
        }
        Compress-Archive @compress
        Remove-Item $subDirectory.fullname -Recurse
    }
}

# Check to make sure all files are of the type html. If not, show a list of the incorrect files to the user and exit. 
function Compare-Filetypes {
    Write-Host "Checking all files are of type .html" -ForegroundColor White -BackgroundColor Green
    $htmlFilesCount = 0
    $nonHtmlFilesCount = 0
    #Powershell's implementation of arrays is stupid. We can modify existing entries, but not add or removed (.IsFixedSize). We have to instead invoke the ArrayList construction.
    $nonHtmlFiles = [System.Collections.ArrayList]@()
    $filesToCheck = @(Get-ChildItem $outputLocation\_working)
    foreach ($filetoCheck in $filesToCheck) {
        $fileSplit = $filetoCheck -split("\.")
        $extension = $fileSplit[1]
        if ($extension -eq "html") {
            $htmlFilesCount++
        }
        else {
            $nonHtmlFiles.Add($filetoCheck)
            $nonHtmlFilesCount++
        }
    }

    if ($htmlFilesCount -ne $filesToCheck.Count){
        Write-Host "There are"$nonHtmlFilesCount" files without the extension .html. Please check input and re-run script." -BackgroundColor Red -ForegroundColor Black 
        foreach ($nonHtmlFile in $nonHtmlFiles) {
            Write-Host $nonHtmlFile -BackgroundColor Red -ForegroundColor Black 
        }
        Pause
        exit
    }
}

function Convert-ImageLanguages {
    # Files will contain references to "_EN.png", and for languages will need changed to the appropriate languages
    Write-Host "Converting image languages in files" -ForegroundColor White -BackgroundColor Green
    $directoryToLocalise = @(Get-ChildItem $outputLocation\_finalOutput -Recurse)
    foreach ($fileToLocalise in $directoryToLocalise) {
        if ($null -eq $fileToLocalise -or ($fileToLocalise.Attributes -band [IO.FileAttributes]::Directory) -ne [IO.FileAttributes]::Directory) {    
            Write-Host $fileToLocalise    
            $nameAndExtension = $fileToLocalise -split("\.")
            $localeLicenceAndLanguage = $nameAndExtension[0] -split("_")
            $language = $localeLicenceAndLanguage[2]
            $licence = $localeLicenceAndLanguage[1]
            # If the licence is GLOBAL (e.g. XXX-YYY_GLOBAL_FR.html), discard licence and only use language portion for image name.
            # Same logic applies if the licence and language match, (e.g. XXX-YYY_GLOBAL_DA_DA.html). This would expect only the language for the name.
            if ($licence -eq "GLOBAL" -or $licence -eq $language) {
                $language = $language.ToUpper()
                $newImageLanguage = "_"+$language+".png"
            } else {
                $licence = $licence.ToUpper()
                $language = $language.ToUpper()
                $newImageLanguage = "_"+$licence+"_"+$language+".png"
            }
            (Get-Content $fileToLocalise.fullName) | Foreach-Object {$_ -replace "_EN.png", $newImageLanguage} | Set-Content $fileToLocalise.fullName -Encoding UTF8         
        }
        else {
           
        }
    } 
}

########################
## Script Entry Point ##
########################

#####################
##  Render Banner  ##
#####################
Clear-Host
Write-Host -BackgroundColor Black -ForegroundColor DarkCyan "===================="
Write-Host -NoNewLine -BackgroundColor DarkMagenta -ForegroundColor Gray "scrmSort"
Write-Host -NoNewLine -BackgroundColor DarkMagenta -ForegroundColor Gray " "
Write-Host -NoNewLine -BackgroundColor DarkMagenta -ForegroundColor Gray " 0.6.4 "
Write-Host -BackgroundColor DarkMagenta -ForegroundColor Gray "    "
Write-Host -BackgroundColor Black -ForegroundColor DarkCyan "===================="
Write-Host -BackgroundColor Red -ForegroundColor White "Please ensure target directory is empty before use. If it is not, the script will delete all existing contents. "

$deliveryLocation = read-host "Please enter the location of the delivery from Translation"
        # If nothing entered, use a default path for testing
        if ($deliveryLocation -eq "" ){
            $deliveryLocation = "C:\Users\finlachr\Documents\_work\StarsCRM Batch Tool\Example delivery"
        }
        if ($deliveryLocation -match ".zip") {
            Write-Host "Zip file found" -ForegroundColor Yellow -BackgroundColor Black
            $deliveryLocation = $deliveryLocation.TrimStart("`"")
            $deliveryLocation = $deliveryLocation.TrimEnd("`"")
            $drive = $deliveryLocation -split(":")
            $decompressedLocation = $drive[0]+":\scrmSort"
            Write-Host $decompressedLocation
            if (Test-Path $decompressedLocation){
                Remove-Item ($decompressedLocation) -Recurse
            }
            try {
                Expand-Archive $deliveryLocation -DestinationPath $decompressedLocation -ErrorAction Stop
            }
            catch {
                Write-Host "An error ocurred:" -BackgroundColor Red -ForegroundColor Black 
                Write-Host $_ -BackgroundColor Red -ForegroundColor Black 
                Write-Host "The script will now exit."
                pause
                exit
            }
            $deliveryLocation = $decompressedLocation
            $fromZip = $true
        }
Write-Host -BackgroundColor DarkCyan -ForegroundColor Black "Copying from" $deliveryLocation
$outputLocation = read-host "Please enter the desired output location. Target location should be an empty folder"
        # If nothing entered, use a default path for testing
        if ($outputLocation -eq "" ){
            $outputLocation = "C:\Users\finlachr\Documents\_upload\_test"
        }
Write-Host -BackgroundColor DarkCyan -ForegroundColor Black "Copying to" $outputLocation
Write-Host -BackgroundColor DarkCyan -ForegroundColor White "Setting up folders and copying delivery"

# Create a working directory so that we don't need to perform any destructive behaviours on the original delivery, and copy files to there. out-null stops the script from echoing info we don't need. If the source and output directories already exist, we remove them to prevent clashes with old files.
if (Test-Path $outputLocation\_deliveryCopy) {
    Remove-Item $outputLocation\_deliveryCopy -Recurse
}
if (Test-Path $outputLocation\_working) {
    Remove-Item $outputLocation\_working -Recurse
}
if (Test-Path $outputLocation\_finalOutput) {
    Remove-Item $outputLocation\_finalOutput -Recurse
}
try {
    mkdir $outputLocation\_deliveryCopy -ErrorAction Stop | out-null 
    mkdir $outputLocation\_working -ErrorAction Stop | Out-Null
    mkdir $outputLocation\_finalOutput -ErrorAction Stop | out-null
}
catch {
    Write-Host "An error occurred trying to create "$outputLocation -BackgroundColor Red -ForegroundColor Black 
    Write-Host $_ -BackgroundColor Red -ForegroundColor Black 
    Write-Host "The script will now exit. Please make sure the specified target area exists and is writable" -BackgroundColor Red -ForegroundColor Black 
    pause
    exit
}

Copy-Item -recurse $deliveryLocation\* $outputLocation\_deliveryCopy

# Get a list of all files and folders inside the deliveryCopy directory, and copy all the files to the working directory (flat) to be analysed and sorted
$copiedCollection = @(Get-ChildItem $outputLocation\_deliveryCopy -Recurse)
$fileCount = 0
foreach ($item in $copiedCollection){
    if ($null -eq $item -or ($item.Attributes -band [IO.FileAttributes]::Directory) -ne [IO.FileAttributes]::Directory) { 
        $fullFilePath=$item.fullname
        Write-Host $fullFilePath
        Copy-Item $fullFilePath $outputLocation\_working
        $filecount++
    }
    else {
        $fullFilePath=$item.fullname
    }
}
Write-Host "Copied" $fileCount "files" -BackgroundColor DarkCyan -ForegroundColor Black

# Get a list of all the files copied to the working location
$workingCollection = @(Get-ChildItem $outputLocation\_working) 
Write-Host -NoNewline "Found"$workingCollection.Count"files in \_working, expected"$fileCount "files. "
if ($workingCollection.Count -eq $fileCount) {
    Write-Host "Moving on!" -BackgroundColor Green -ForegroundColor Black
}
else {
    Write-Host ""
    Write-Host "A different amount of files exists in the _working directory than were found in the original delivery. " -BackgroundColor Red -ForegroundColor Black 
    Write-Host "The script will exit. Please check the delivery for duplicated filenames." -BackgroundColor Red -ForegroundColor Black
    pause
    exit
}

foreach ($fileToCheck in $workingCollection){
    # Translation submit Slovnia files as si.html, but StarsCRM expects sl.html instead. 
    $sloveniaPattern = "si.html"
    if($fileToCheck.Name -match $sloveniaPattern){ 
        Write-Host "File named si.html found. Renaming to sl.html" -ForegroundColor Yellow -BackgroundColor Black
        $new_name = $fileToCheck.Name -replace $sloveniaPattern, "sl.html"
        Rename-Item -Path $fileToCheck.FullName -NewName $new_name
    }
}

Compare-Filetypes

# Repopulate the array of files with any modified names
$workingCollection = @(Get-ChildItem $outputLocation\_working) 
foreach ($file in $workingCollection) {
        Copy-Languages $file
}
Convert-ImageLanguages
Compress-Results

Write-Host "File copying and sorting complete." -BackgroundColor Green -ForegroundColor White
Write-Host "Please check the output in "$outputLocation\_finalOutput". (Press enter to open folder)" -BackgroundColor DarkCyan -ForegroundColor Black 
Pause
Invoke-Item $outputLocation\_finalOutput
Write-Host "All other folders will be cleaned and removed. Press enter to confirm, or exit script if you wish to keep these files." -BackgroundColor DarkRed -ForegroundColor White 
Pause
Remove-Item $outputLocation\_deliveryCopy -Recurse
Remove-Item $outputLocation\_working -Recurse
if ($fromZip) {
    Remove-Item $decompressedLocation -Recurse
}
Write-Host "Folders cleaned and ready for import into StarsCRM." -BackgroundColor Green -ForegroundColor White
Pause

# Changelog
# v0.1.0 - Initial release
# v0.1.1 - Added support for renaming Slovenia files to match the sl.html format expected by SCRM
# v0.2.1 - Added proper error handling when attempting to create the output directories
# v0.3.1 - Added support for passing in ZIP archives as the initial delivery
# v0.4.1 - Added check to ensure all copied files are of type "html"
# v0.5.1 - Logic updated to accept new naming convention
# v0.6.1 - Added function to localise image names within delivery
# v0.6.2 - Script will now open the _finalOutput folder for you once the archive compression is complete
# v0.6.3 - Fixed bug where image names were returned as ".PNG" instead of ".png"
# v0.6.4 - Fixed bug where cyrillic characters were mangled when updating image names