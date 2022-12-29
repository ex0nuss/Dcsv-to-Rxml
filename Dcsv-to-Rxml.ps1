<#
==================================================================
Titel: Dcsv-to-Rxml (DenonCSV-to-RekordboxXML)
Version: 1.1
Date: 2022-12-28
==================================================================
Parameter: 
    - InputCsv: Uses the given CSV file to convert it to a rekordbox.xml
      --> If empty it uses all CSVs in the $PSScriptRoot 
    - OutputXml: Destination file for the rekordbox.xml
      --> If empty it generates a rekordbox.xml in the $PSScriptRoot
Description:
    - Converts one or more Engine DJ playlist CSV files into the rekordbox.xml format
---
Author: ex0nuss
==================================================================
Change log:
    Version 1.0:
        - Created this Script
    Version 1.1:
        - Added Validate-ObjectHeader for the CSV
#> 



### Input params ###
param (
    [Parameter(Position=1)][string]$InputCsv,
    [Parameter(Position=2)][string]$OutputXml
)



### Functions ###
function Read-HostDefaultValue {
    Param(
        [parameter(Mandatory=$true)][String]$Promt,
        [parameter(Mandatory=$true)][String]$DefaultValue,
        [parameter(Mandatory=$true)]$PossibleValues
    ) 

    # Checks if $DefaultValue is NOT in $PossibleValues
    if ($PossibleValues -notcontains $DefaultValue) {
        throw "DefaultValue must be in PossibleValues"
    }

    # Makes the DefaultValue upper case in PossibleVaules
    $PossibleValues_Default = $PossibleValues.Replace($DefaultValue,$DefaultValue.ToUpper())
    # Build the promt --> Promt [a/B/c]
    $promtCombined = $($promt + ' [' + $($PossibleValues_Default -join '/') + ']')
    
    $retry = $false

    do {
        # Checks if it already looped
        if ($retry) {
            Write-Host -ForegroundColor Yellow "Please input a vaild value:"
            $PossibleValues | ForEach-Object { Write-Host -ForegroundColor Yellow "`t- $_" }
        }
        # Sets retry to true for seccond interation
        $retry = $true

        # Read input
        $userInput = Read-Host -Prompt "$promtCombined"

        # Tests if input is empty --> default is used
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            $userInput = $DefaultValue
        }
        
    } while ($PossibleValues -notcontains "$userInput")

    return $userInput
}
# Tests if object / array has all the required headers
function Validate-ObjectHeader {
    Param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]$InputToTest,
        [parameter(Mandatory=$true)]$RequiredHeaders,
        [parameter(Mandatory=$false)][switch]$CheckIfValueIsNotNull = $false
    )

    # Loops through input
    foreach ($obj in $InputToTest) {

        # Tests for every required header
        foreach ($RequiredHeader in $RequiredHeaders) {

            # Tests if $obj has the RequiredHeader
            if ($obj.PSobject.Properties.Name -notcontains $RequiredHeader) {
                return $false

            # If $obj has all the required header
            } else {
                # if not null value should be checked
                if ($CheckIfValueIsNotNull) {

                    # Gets the value from the RequiredHeader
                    $RequiredHeaderValue =  $obj | Select-Object -ExpandProperty $RequiredHeader
                    # Checks if it nut null or empty
                    if ([string]::IsNullOrWhiteSpace($RequiredHeaderValue)) {
                        return $false
                    }
                }
            }
        }
    }

    return $true
}



### Parameter handling: InputCsv ###
# If input is empty 
if ([string]::IsNullOrEmpty($InputCsv)) {

    Write-Host -ForegroundColor Yellow "No InputCsv given, using `"$PSScriptRoot`""
    $CsvFiles = Get-ChildItem -Path "$PSScriptRoot\*" -Include *.csv

    Write-Host "All CSVs will be used:" 
    $CsvFiles | Select-Object -ExpandProperty 'Name' | ForEach-Object { Write-Host "`t- $_" }
    $inputContinue = Read-HostDefaultValue -Promt "Do you want to continue" -DefaultValue 'n' -PossibleValues 'y','n'
    if ($inputContinue -eq 'n') { throw 'Aborted by user.'}

# If NOT empty
} else {
    $CsvFiles = Get-ChildItem "$InputCsv"
}



### Parameter handling: OutputXml ###
# If output is empty 
if ([string]::IsNullOrEmpty($OutputXml)) {

    # Uses dir from input
    # ToDo: Use folder when InputCsv is empty
    $OutputFile = "$($CsvFiles[0].Directory)\rekordbox.xml"
    Write-Host -ForegroundColor Yellow "No OutputXml given, using input directory: `"$OutputFile`""

# If NOT empty
} else {
    $OutputFile = $OutputXml
}



### Vars ###
# Csv with all playlists
$Csv = $null
# Global ID (counter) for playlist links
$ID = 0



### Generating global OBJ with all playlists ###
# Lopps through all playlist CSVs to get one big OBJ
Foreach ($CsvFile in $CsvFiles) {

    # Imports CSV as file for replacing
    $Cont = Get-Content "$($CsvFile.FullName)" -Encoding UTF8

    # Replaces all , with ; expect in double quotes (only delimiters)
    $CsvSmall = $Cont -replace ',+(?=([^"]*"[^"]*")*[^"]*$)', ';'
    # Replacments for reading the csv correctly
    $CsvSmall = $CsvSmall.replace('File name','Filename')
    $CsvSmall = $CsvSmall.replace('#','Number')

    # Converts CSV to Objs
    $CsvSmall = $CsvSmall | ConvertFrom-Csv -Delimiter ';' 
    # Adds Playlist- and ID-property to Obj
    $CsvSmall | ForEach-Object { 
        Add-Member -InputObject $_ -NotePropertyName 'Playlist' -NotePropertyValue $($CsvFile.BaseName); 
        Add-Member -InputObject $_ -NotePropertyName 'ID' -NotePropertyValue $ID; 
        $ID++ }
    # Checks if object headers are missing
    $HeadersComplete = Validate-ObjectHeader -InputToTest $CsvSmall -RequiredHeaders 'Title','Artist','Filename' -CheckIfValueIsNotNull
    if ($HeadersComplete -eq $false) {
        throw "Some headers in the CSV are missing or are empty, please make sure the following headers are there: Title,Artist,Filename"
    } 
    $Csv = $Csv + $CsvSmall
}


### XML header ### 
$xmlWriter = New-Object System.XMl.XmlTextWriter("$OutputFile",[System.Text.Encoding]::UTF8)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"

$xmlWriter.WriteStartDocument()

$xmlWriter.WriteStartElement("DJ_PLAYLISTS")
    $xmlWriter.WriteAttributeString("Version","1.0.0")

    $xmlWriter.WriteStartElement("PRODUCT")
        $xmlWriter.WriteAttributeString("Name","Dcsv-to-Rxml (DenonCSV-to-RekordboxXML)")
        $xmlWriter.WriteAttributeString("Version","1.0.0")
        $xmlWriter.WriteAttributeString("Company","ex0nuss")
    $xmlWriter.WriteEndElement()

    $xmlWriter.WriteStartElement("COLLECTION")
        $xmlWriter.WriteAttributeString("Entries","$($Csv.Count)")


### Dynamically adding tracks to XML ###
Write-Host "Adding tracks to XML"
foreach ($track in $Csv) { 
    Write-Host "`t- Working on: $($track.ID). $($track.Artist) - $($track.Title)"

    # Get file extension
    $GciTrack  = Get-Childitem $($track.Filename)
    $Kind = $GciTrack | Select-Object -ExpandProperty 'Extension'
    $Kind = $Kind.Replace('.','')
    $Kind = $Kind.ToUpper()
    # Get date added
    $DateAdded = $GciTrack | Select-Object -ExpandProperty 'CreationTime'
    $DateAdded = Get-Date $DateAdded -Format 'yyyy-MM-dd'

    # Add track in XML
    $xmlWriter.WriteStartElement("TRACK")
        $xmlWriter.WriteAttributeString("TrackID","$($track.ID)")
        $xmlWriter.WriteAttributeString("Name","$($track.Title)")
        $xmlWriter.WriteAttributeString("Artist","$($track.Artist)")
        $xmlWriter.WriteAttributeString("Album","$($track.Album)")
        $xmlWriter.WriteAttributeString("Genre","$($track.Genre)")
        $xmlWriter.WriteAttributeString("Kind","$Kind")
        $xmlWriter.WriteAttributeString("Location","file://localhost/$($track.Filename)")
        $xmlWriter.WriteAttributeString("DateAdded","$DateAdded")
    $xmlWriter.WriteEndElement()
}
$xmlWriter.WriteEndElement()


### Dynamically adding tracks to playlist in XML ###
$xmlWriter.WriteStartElement("PLAYLISTS")
    
    # Gets all possible playlists
    $PlaylistNames = $Csv | Select-Object -ExpandProperty 'Playlist' -Unique

    # Playlists header
    $xmlWriter.WriteStartElement("NODE")
    $xmlWriter.WriteAttributeString("Type","0")
    $xmlWriter.WriteAttributeString("Name","ROOT")
    $xmlWriter.WriteAttributeString("Count","$($PlaylistNames.Length)")

    # Loops through all playlists
    Foreach ($PlaylistName in $PlaylistNames) {

        # Gets only the tracks from each playlist 
        $PlaylistTracks = $Csv | Where-Object -Property 'Playlist' -EQ "$PlaylistName"

        # Header for individual playlist
        $xmlWriter.WriteStartElement("NODE")
            $xmlWriter.WriteAttributeString("Name","$PlaylistName")
            $xmlWriter.WriteAttributeString("Type","1")
            $xmlWriter.WriteAttributeString("KeyType","0")
            $xmlWriter.WriteAttributeString("Entries","$($PlaylistTracks.Count)")

        Write-Host "Adding tracks to playlist $PlaylistName in XML"

        # Foreach entry in playlist
        foreach ($entry in $PlaylistTracks) {
            Write-Host "`t- Adding $($entry.ID)"

            # Adding track to playlist by ID
            $xmlWriter.WriteStartElement("TRACK")
                $xmlWriter.WriteAttributeString("Key","$($entry.ID)")
            $xmlWriter.WriteEndElement()
        }

        $xmlWriter.WriteEndElement()
    }

    $xmlWriter.WriteEndElement()
$xmlWriter.WriteEndElement()


# XML fotter
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()


<#
Copy-Item -Force -Verbose -Path C:\Users\Max\Downloads\Wallis.xml "C:\Users\Max\AppData\Roaming\Pioneer\rekordbox\rekordbox.xml"
#>

<#
### Example XML only code (static) ###

$xmlWriter = New-Object System.XMl.XmlTextWriter("C:\Users\Max\Downloads\Wallis.xml",[System.Text.Encoding]::UTF8)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"

$xmlWriter.WriteStartDocument()

$xmlWriter.WriteStartElement("DJ_PLAYLISTS")
    $xmlWriter.WriteAttributeString("Version","1.0.0")

    $xmlWriter.WriteStartElement("PRODUCT")
        $xmlWriter.WriteAttributeString("Name","Dcsv-to-Rxml")
        $xmlWriter.WriteAttributeString("Version","0.0.1")
        $xmlWriter.WriteAttributeString("Company","DenonCSV-to-RekordboxXML")
    $xmlWriter.WriteEndElement()

    $xmlWriter.WriteStartElement("COLLECTION")
        $xmlWriter.WriteAttributeString("Entries","1")

        $xmlWriter.WriteStartElement("TRACK")
            $xmlWriter.WriteAttributeString("TrackID","0")
            $xmlWriter.WriteAttributeString("Name","")
            $xmlWriter.WriteAttributeString("Artist","")
            $xmlWriter.WriteAttributeString("Album","")
            $xmlWriter.WriteAttributeString("Genre","")
            $xmlWriter.WriteAttributeString("Kind","")
            $xmlWriter.WriteAttributeString("Location","file://localhost/")
            $xmlWriter.WriteAttributeString("DateAdded","")
            #$xmlWriter.WriteAttributeString("Size","")
            #$xmlWriter.WriteAttributeString("TotalTime","")
            #$xmlWriter.WriteAttributeString("DiscNumber","")
            #$xmlWriter.WriteAttributeString("TrackNumber","")
            #$xmlWriter.WriteAttributeString("Year","")
            #$xmlWriter.WriteAttributeString("DateAdded","")
            #$xmlWriter.WriteAttributeString("BitRate","")
            #$xmlWriter.WriteAttributeString("SampleRate","")
            #$xmlWriter.WriteAttributeString("Comments","")
            #$xmlWriter.WriteAttributeString("Rating","")
            #$xmlWriter.WriteAttributeString("Remixer","")
            #$xmlWriter.WriteAttributeString("Tonality","")
            #$xmlWriter.WriteAttributeString("Label","")
        $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndElement()

    $xmlWriter.WriteStartElement("PLAYLISTS")
        $xmlWriter.WriteStartElement("NODE")
        $xmlWriter.WriteAttributeString("Type","0")
        $xmlWriter.WriteAttributeString("Name","ROOT")
        $xmlWriter.WriteAttributeString("Count","1")

        $xmlWriter.WriteStartElement("NODE")
            $xmlWriter.WriteAttributeString("Name","Wallis (1)")
            $xmlWriter.WriteAttributeString("Type","1")
            $xmlWriter.WriteAttributeString("KeyType","0")
            $xmlWriter.WriteAttributeString("Entries","1")

            $xmlWriter.WriteStartElement("TRACK")
                $xmlWriter.WriteAttributeString("Key","0")
            $xmlWriter.WriteEndElement()

        $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndElement()
$xmlWriter.WriteEndElement()


$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()
#>