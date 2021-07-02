<#
.SYNOPSIS
    A script to import class instances (work item or configuration item based classes), link related items and import file attachments exported using the Export-ClassInstances.ps1 script (https://github.com/bennyguk/Export-ClassInstances)
.DESCRIPTION
    This script could be useful if you need to import class instances in bulk when makeing changes to a custom class
    that are not upgrade compatible and have first been exported using Export-ClassInstances.ps1  (https://github.com/bennyguk/Export-ClassInstances).
    
    For more information, please see https://github.com/bennyguk/Import-ClassInstances

.PARAMETER ClassName
    Specifies the class name you wish to work with.

.PARAMETER FilePath
    Specifies the path to the folder you wish to import class instances and file attachments from.

.PARAMETER FileName
    Specifies name of the CSV file - Will default to Export.csv

.PARAMETER ComputerName
    Specifies the SCSM server to connect to - Will default to localhost.
    
.EXAMPLE
    Import-ClassInstances.ps1 -ClassName MyClass -FilePath c:\MyClassImport -FileName MyClassImport.csv -ComputerName MySCSMServer
#>
Param (
    [parameter(mandatory)][string] $ClassName,
    [parameter(Mandatory, HelpMessage = "Enter a path to the exported CSV directory, excluding the filename")][string] $FilePath,
    [string] $FileName = "Export.csv",
    [parameter(HelpMessage = "Enter Managment Server Computer Name")]
    [string] $ComputerName = "localhost"
)

# Import modules and set SCSM Management Server computer name
$smdefaultcomputer = $ComputerName
Import-Module SMlets
$managementGroup = new-object Microsoft.EnterpriseManagement.EnterpriseManagementGroup "$smdefaultcomputer"

# Get the class information from Service Manager
$class = Get-SCSMClass | Where-Object { $_.Name -eq $className }

# Check to see if the class exists
if (!$class) {
    Write-Host "Could not load class '$className'. Please check the class name and try again."
    Exit
}
# Check to see if the file path exists
If (!(Test-Path $FilePath)) {
    Write-Host "Could not find '$FilePath'. Please check the path name and try again."
    Exit
}

# Import CSV files
if (Test-Path $FilePath\$FileName) {
    $ImportCsv = Import-Csv -Path $FilePath\$FileName
    if ($FileName -match ".") {
        $splitName = $FileName.Split('.')
        $relFileName = "$($splitName[0])-relationships.$($splitName[1])"
    }
    else {
        $relFileName = "$FileName-relationships"
    }
}
else {
    Write-Host "Could not find $FilePath\$FileName. Please check the path and the filename and try again"
    Exit
}
if (Test-Path $FilePath\$relFileName) {
    $ImportRelCsv = Import-Csv -Path $FilePath\$relFileName
}
else {
    Write-Host "Could not find $FilePath\$relFileName. Please check the path and the filename and try again"
    Exit
}

# Get relationship class for file attachments for the class we are working with, either configuration Item or Work Item
if ($Class.GetBaseTypes().Name -contains "System.ConfigItem") {
    $fileAttachmentRel = Get-SCSMRelationshipClass "System.ConfigItemHasFileAttachment"
}
else {
    $fileAttachmentRel = Get-SCSMRelationshipClass "System.WorkItemHasFileAttachment"
}

# Get type projection inforamtion for the file attachment
$tp = Get-SCSMTypeProjection $class
# Create a hashtable with the any type projections for the class that have a name of System.FileAttachment
$tpOptions = @{}
$i = 0
foreach ($tpClassName in $tp) {
    If ($tpClassName.TypeProjection.Key.Class.name -match "System.FileAttachment") {
        $tpOptions[$i] += $tpClassName.TypeProjection.Name
        $i++
    }
}
# If there is more than one type projection, display the option to select the most appropriate 
if ($tpOptions.Count -gt 1) {
    $t1 = @()
    Write-Host ("The class '$class' appears to have more than one type projection for file attachments. Please select which type projection you wish to use:`n`r") -ForegroundColor Green
    $propertyCount = $tpOptions.count
    for ($i2 = 0 ; $i2 -lt $propertyCount ; $i2++) {
        Write-Host $i2 "-" $tpOptions[$i2] "`n`r" -ForegroundColor Green
        $t1 += $i2
    }
    $a = $t1 -join "|"

    do {
        Write-Host "Please enter a value between 0 and" ($i2 - 1)": " -ForegroundColor Green -NoNewline
        $typeProjNumber = Read-Host 
    }
    while ($typeProjNumber -notmatch $a) {
    }
    $classTypeProj = $tpOptions[$typeProjNumber -as [int]]
}
else {
    $classTypeProj = $tp.TypeProjection.Name
}
Function Add-FileAttachment {

    param($files, $documentID)
    $filecounter = 0
    Foreach ($file in $files) {
        $fileCounter++
        Write-Progress -Id 1 -ParentId 0 -Status "Processing $($FileCounter) of $($files.count)" -Activity "Importing files" -CurrentOperation $file.Name -PercentComplete (($fileCounter / $files.count) * 100)
        $fileClass = Get-SCSMClass -name "System.FileAttachment"
        $typeProj = Get-SCSMObjectProjection $classTypeProj -filter "ID -eq $documentID"
        $fileMode = [System.IO.FileMode]::Open
        $read = new-object System.IO.FileStream $file.FullName, $fileMode
        $newFileAttach = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($managementGroup, $fileClass)

        $newFileAttach.Item($fileClass, "Id").Value = [Guid]::NewGuid().ToString()
        $newFileAttach.Item($fileClass, "DisplayName").Value = $file.Name
        $newFileAttach.Item($fileClass, "Description").Value = $file.Name
        $newFileAttach.Item($fileClass, "Extension").Value = $file.Extension
        $newFileAttach.Item($fileClass, "Size").Value = $length
        $newFileAttach.Item($fileClass, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
        $newFileAttach.Item($fileClass, "Content").Value = $read
        $typeProj.__base.Add($newFileAttach, $fileAttachmentRel.Target)
        $typeProj.__base.Commit()

        $read.close()
    }
}

# Main script

# Create hashtables from the CSVs
$htClassInstance = @{}
$htRelInstance = @{}
if ($ImportCsv[0]) {
    $cICounter = 0
    Foreach ($classInstance in $ImportCsv) {
        $cICounter++
        Write-Progress -Id 0 -Status "Processing $($cICounter) of $($ImportCsv.count)" -Activity "Importing all instances of $($class.DisplayName)" -CurrentOperation $classInstance.DisplayName -PercentComplete (($cICounter / $ImportCsv.count) * 100)
        # Add all keys and values to a hashtable and set any blank values to null where they exist. This prevents errors about costing to string later.
        $propertyCount = $classInstance.psobject.properties.name.count
        for ($i = 0 ; $i -lt $propertyCount ; $i++) {
            $htClassInstance[$classInstance.psobject.properties.name[$i]] = $classInstance.psobject.properties.value[$i]
            if ($htClassInstance[$classInstance.psobject.properties.name[$i]] -eq "") {
                $htClassInstance[$classInstance.psobject.properties.name[$i]] = $null
            }
        }
        # Create a new instance of the target class with the keys and values defined in the hashtable
        Try {
            $newClassInstance = New-SCSMObject -Class $Class -PropertyHashtable $htClassInstance -PassThru -ErrorAction Stop
        }
        catch {
            Write-Host ("An error has occured creating a new class instance. The error message was:") -ForegroundColor Red
            Write-Host $_ -ForegroundColor Red
            Exit
        }
        # Attach all exported file attachments if they exists
        $docID = $newClassInstance.Id
        if (Test-Path $FilePath\ExportedAttachments\$docID) {
            $files = get-childitem $FilePath\ExportedAttachments\$docID
            Add-FileAttachment -files $files -documentID $docID
        }
        # Create relationships and related items if they exist
        if ($ImportRelCsv[0]) {

            # Get the key property of the newly created class instance
            $ciKey = ($newClassInstance.GetProperties() | Where-Object { $_.Key -eq "True" }).Name

            # If the class does not have a key property, add the internal ID instead. This is needed incase you change your class to have a key property later, otherwise the import will fail. The ID or key property is also used to import files.
            if (!$ciKey) {
                $ciKey = "ID"
            }

            # Load the related items for only the key property that corresponds to the $newClassInstance key property
            $relInstance = $ImportRelCsv | Where-Object { $_.$ciKey -eq $newClassInstance.$ciKey }
            # Add all keys and values to a hashtable and set any blank values to null where they exist.
            $propertyCount2 = $relInstance.psobject.properties.name.count
            for ($i2 = 0 ; $i2 -lt $propertyCount2 ; $i2++) {
                $htRelInstance[$relInstance.psobject.properties.name[$i2]] = $relInstance.psobject.properties.value[$i2]
                if ($htRelInstance[$relInstance.psobject.properties.name[$i2]] -eq "") {
                    $htRelInstance[$relInstance.psobject.properties.name[$i2]] = $null
                }
            }
            # Create new relationship instances from the $htRelInstance hashtable
            foreach ($relationshipObject in $htRelInstance.GetEnumerator()) {

                # filter out file attachment related relationships as these are handled by the Add-Attachment function.
                if ($relationshipObject.Key -notlike "*Attachment*") {
                    $relationship = Get-SCSMRelationshipClass $relationshipObject.Key
                    if ($relationshipObject.Value) {
                        if ($relationshipObject.Value -match ",") {
                            $relationshipObjectValues = $relationshipObject.Value -split ","
                            foreach ($relationshipObjectValue in $relationshipObjectValues) {
                                # sometimes there maybe more than one related item of a particular display name, so the script will just choose the first.
                                $relItemValue = (Get-SCClassInstance -Class ($relationship.Target.Class) -Filter "DisplayName -eq $relationshipObjectValue") | Select-Object -first 1
                                if ($relItemValue) {
                                    New-SCRelationshipInstance -RelationshipClass $relationship -Source $newClassInstance -Target $relItemValue -PassThru > $null
                                }
                            }

                        }
                        else {
                            $relValueName = $relationshipObject.Value
                            # sometimes there maybe more than one related item of a particular display name, so the script will just choose the first.
                            $relItemValue = (Get-SCClassInstance -Class ($relationship.Target.Class) -Filter "DisplayName -eq $relValueName") | Select-Object -First 1
                            if ($relItemValue) {
                                New-SCRelationshipInstance -RelationshipClass $relationship -Source $newClassInstance -Target $relItemValue -PassThru > $null
                            }
                        }
                    }    
                }
            }
        }
        else {
            Write-Host "No related items to process. Please check that the relationships csv file has data and try again"
            exit
        }
    }
}
else {
    Write-Host "Please check that the csv has data and try again"
    exit
}
