##############################
### Single Image Downloader###
###                        ###
### Revised: 2020/09/20    ###
### Author: Zachary Eagle  ###
### Tested with:           ###
###   PowerShell 5.1       ###
###   Win 10 (1803) Ent    ###
###   DCMTK 3.6.4          ###
##############################

###########################  Read Me #######################################
### You can use this tool to download a specific image from a specific series from a specifc study
### You will need to have DCMTK installed on your computer
### Your computer will need to have query and retrieve access from PACS
### Modify the parameters below to fit your project
#############################################################################

##############################
####    Parameter Files   ####
##############################
. ".\PACS_info.ps1"
. ".\Project_info.ps1"

##############################
#### Script - Do not edit ####
##############################
function Try-Command($Command, $Parameters, $queryfile) {
     $i = 0
     Write-Host $command $parameters
     do{
        $output = & $Command $Parameters $queryfile
        #Start-Sleep -Seconds 15
        $i++
        }until(($output | Select-String -Pattern "0x0000: Success" -Quiet) -or ($i -eq 20))
        
        if(($output | Select-String -Pattern "0x0000: Success" -Quiet)){
            if($queryfile){#Remove-Item $queryfile}
            }
            }
        else{
            $error = ""
            $error = "Error on command: "+$command+$parameters+$queryfile
            Add-Content $errorLog -Value $error
            return "Error"
        }
 }

function Find-DICOM($Query){
    write-host "Finding"
    $findexe = "C:\Program Files\dcmtk\bin\findscu.exe" 
    $findparam = @(
        '-S',
        '-d',
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT)
    $findParameters = $findparam + $Query
    Try-Command $findexe $findParameters
}

function Move-DICOM($Query, $QueryFile){
    Write-Host "Moving"
    $moveexe = "C:\Program Files\dcmtk\bin\movescu.exe" 
    $param = @(
        '-S',
        '-d',
        '+xa',
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT,
        '--port', $myPort)
    $moveParameters = $param + $Query
    Try-Command $moveexe $moveParameters $QueryFile
}

function Create-UniqueDirectory($Path, $Name){
        $i = ''
        $name = $name  -replace '[^A-Za-z0-9_.]', ''
        $fullPath = $path +"\"+$name
        while(Test-Path ($fullPath+$i)){
            $i = $i+1 -as [int]
            }
        $folder = New-Item -ItemType directory -Path ($fullPath+$i)
        return $folder.FullName
}

$ProjectPath = $ProjectFolder +"\"+ $ProjectName
$errorLog = $projectPath+ "errors.txt"
#Check if path exists, if it does, error, if not create it.
If(Test-Path $ProjectPath -PathType Container){
    Throw "Project already exists. Choose a different name"
    }Else{
    New-Item -ItemType directory -Path $ProjectPath
    }
If(-not(Test-Path $AccList -PathType leaf)){
    Throw "File does not exist, check path"
    }
#Create a folder for each accession and get the study level query file
$accessions = Import-Csv $AccList
$find = $ProjectPath + "\Find\"
$out = $ProjectPath + "\Output\"

$findOutput = New-Item -ItemType directory -Path $find
$dicomOutput = New-Item -ItemType directory -Path $out
ForEach ($accession in $accessions){
    $findParam = @(
            '-k', 'QueryRetrieveLevel=STUDY',
            '-k', 'SOPInstanceUID'
            '-k', 'SeriesInstanceUID'
            '-k', 'StudyInstanceUID'
            '-k', -join ('AccessionNumber=', $accession.AccessionNumber),
            '-k', -join ('SeriesNumber=', $accession.SeriesNumber),
            '-k', -join ('InstanceNumber=', $accession.InstanceNumber),
            '-X',
            '-od', $findOutput
        )
    Find-DICOM $findParam

   
    $rsps = get-childitem $findOutput
    ForEach ($rsp in $rsps){
        $moveParam = @(
        $rsp.FullName,
        '-k', 'QueryRetrieveLevel=IMAGE',
        '-od', $dicomOutput.FullName)

        Move-DICOM $moveParam}

}
$outputs = Get-ChildItem $dicomOutput
foreach ($file in $outputs){
    $newname = $file.name + ".dcm"
    Rename-Item -Path $file.FullName -newname $newname
    }