$Config = Import-PowerShellDataFile -Path ".\Config.psd1"

$7Zip = $Config.PathTo7Zip
$ZIPPath = $Config.PathToZIPFiles
$ZIPExtractPath = $Config.PathExtract
$7ZipPass = $Config.PasswordZipFiles

$PSTFiles = Get-ChildItem -Filter *.zip -Path $ZIPPath -Recurse

$Count = 0
ForEach($PSTFile in $PSTFiles)
{
    $Count ++
    Write-Progress -Activity "Working on Unzipping files" -PercentComplete (($Count/$PSTFiles.Count)*100)
    Write-Host "[$(Get-Date -format hh:mm:ss)] Unzipping $($PSTFile.FullName)"
    $Command = "& $7Zip e $($PSTFile.Fullname) -o$ZIPExtractPath -y -tzip -p$7ZipPass"
    Invoke-Expression $Command
}
Write-Progress -Activity "Working on Unzipping files" -Completed


$global:O365DataFile = @()
Function WriteToO365DataFileArray
{
    Param($PSTName,$Mailbox)

    $WriteArray = $NULL
    $WriteArray = New-Object PSObject
    $WriteArray | Add-Member -MemberType NoteProperty -Name "Workload" -Value "Exchange"
    $WriteArray | Add-Member -MemberType NoteProperty -Name "FilePath" -Value $Null
    $WriteArray | Add-Member -MemberType NoteProperty -Name "Name" -Value $PSTName
    $WriteArray | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $Mailbox
    $WriteArray | Add-Member -MemberType NoteProperty -Name "IsArchive" -Value $False
    $WriteArray | Add-Member -MemberType NoteProperty -Name "TargetRootFolder" -Value "/Email Archive" #Use / for importing into Root folder
    $WriteArray | Add-Member -MemberType NoteProperty -Name "ContentCodePage" -Value $Null
    $WriteArray | Add-Member -MemberType NoteProperty -Name "SPFileContainer" -Value $Null
    $WriteArray | Add-Member -MemberType NoteProperty -Name "SPManifestContainer" -Value $Null
    $WriteArray | Add-Member -MemberType NoteProperty -Name "SPSiteUrl" -Value $Null
    $global:O365DataFile += $WriteArray
}


$O365PSTs = Get-ChildItem -Filter *.pst -Path $ZIPExtractPath

ForEach($O365PST in $O365PSTs)
{
    $Mailbox = $($O365PST.Name) -replace "_([0-9]+)\.pst",""
    Write-Host $Mailbox
    WriteToO365DataFileArray -PSTName $($O365PST.Name) `
                            -Mailbox $Mailbox
}

$global:O365DataFile | Export-CSV -Path "C:\Users\cirvinadmin1\desktop\O365Imports.csv" -NoTypeInformation