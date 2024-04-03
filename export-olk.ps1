if ($args.count -eq 0 -or $args.count -gt 2){
    printUsage    
}
$inputFile=$args[0]
if ($args.count -eq 2){
    $outputFile=$args[1]
}
else {$outputFile=$null}

function printUsage(){
    write-host "Usage: export-olk.ps1 <inputfile> [outputfile]
    
    Inputfile must be an olk15RecentAddresses file. 
    Outputfile will be CSV formatted.  Outputfile may be omitted to write to stdout.
    If outputfile is omitted, the stdout will pipe objects instead of CSV.

    Example: .\export-olk.ps1 samplefile.olk15RecentAddresses outputfile.csv
    "
}
$fileBytes=[System.IO.File]::ReadAllBytes($inputFile)
$startMail=66
$endMail=$null
$mailIndices=@()
$nameIndices=@()

$i=$startMail
do {
    $i=$i+1
} until ($filebytes[$i] -eq 0 -and $fileBytes[$i+1] -eq 0)
$endMail=$i
$startMailIndex=$i+2
$i=$startMailIndex
$index=$startMail
do {
    $mailIndices+=$index
    $indexOffset=([int]$fileBytes[$i+3] -shl 8)+$fileBytes[$i+2]
    $index=$startMail+$indexOffset
    $i=$i+4
} until ($index -ge $endmail)
$mailIndices+=$endMail
$endMailIndex=$i

$addresses=@()
for ($i=0;$i -lt ($mailIndices.Count -1); $i++){
    $strBytes=for($j=$mailIndices[$i];$j -lt $mailIndices[$i+1];$j++){
        $fileBytes[$j]
    }
    $addresses+=[PSCustomObject]@{
        Type="SMTP";
        Name="";
        Address = [System.Text.Encoding]::UTF8.GetString($strBytes);        
    }
}

$startNames=$endMailIndex+2
$i=$startNames
do {
    $i=$i+1
} until ($filebytes[$i] -eq 0 -and $fileBytes[$i+1] -eq 0)
$endNames=$i

$startNameIndex=$endNames+3
$i=$startNameIndex
$index=$startNames
do{
    $nameIndices+=$index
    $indexOffset=([int]$fileBytes[$i+3] -shl 8)+$fileBytes[$i+2]
    $index=$startNames+$indexOffset
    $i=$i+4
} until ($index -ge $endNames)
$nameIndices+=$endNames+1
$endNameIndex=$i

for ($i=0;$i -lt ($nameIndices.Count -1); $i++){
    if ($nameIndices[$i] -eq $nameIndices[$i+1]){
        continue
    }
    $strBytes=for($j=$nameIndices[$i];$j -lt $nameIndices[$i+1];$j++){
        $fileBytes[$j]
    }

    $addresses[$i].Name=[System.Text.Encoding]::Unicode.GetString($strBytes)
}

if ($outputFile) {
    $addresses | export-csv -NoTypeInformation -Path "$outputFile"
    write-host "Exported CSV to $outputFile"
}
else {$addresses}