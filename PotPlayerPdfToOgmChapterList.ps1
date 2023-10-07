
param (
    [string]$sourceFilePath = "",
    [string]$locale = "zh-TW"
)

#Get current path from script file.
$CurrentPS1File = $(Get-Item -Path "$PSCommandPath")
Set-Location "$($CurrentPS1File.PSParentPath)" | Out-Null

#Check if that has pbf file from param.
$Result = ''

if (-not ([string]::IsNullOrEmpty($sourceFilePath))) {
    if ( $(Test-Path -Path $sourceFilePath -PathType Leaf) -eq $false ) {
        [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
        $sourceFileItem = New-Object System.Windows.Forms.OpenFileDialog
        $sourceFileItem.Filter = "PotPlayer Bookmark Files (*.pbf)|*.pbf" 
        If($sourceFileItem.ShowDialog() -eq "OK") {
            $sourceFilePath = $sourceFileItem.FileName
            $sourceFileItem = Get-Item $sourceFilePath
            $Result_Target = $('來源檔案："' + $sourceFileItem.FullName + '"')
            $Result = $Result + "`r`n"  + $Result_Target
            Write-Output $Result
        }
    }
} else {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    $sourceFileItem = New-Object System.Windows.Forms.OpenFileDialog
    $sourceFileItem.Filter = "PotPlayer Bookmark Files (*.pbf)|*.pbf"
    If($sourceFileItem.ShowDialog() -eq "OK") {
        $sourceFilePath = $sourceFileItem.FileName
        $sourceFileItem = Get-Item $sourceFilePath
        $Result_Target = $('來源檔案："' + $sourceFileItem.FullName + '"')
        $Result = $Result + "`r`n"  + $Result_Target
        Write-Output $Result
    }
}

#Set Ogm file path same as source file.
$OutputPath = $(Get-Item $sourceFileItem.PSParentPath).FullName + '\' + $($sourceFileItem.BaseName) + '-Chapters-' + $(Get-Date).ToString('yyyyMMdd-HHmmss') + '.txt'

#Parsing pbf file as object.
$SrcObj = Get-Content -Path "$($sourceFileItem.FullName)" -Raw -Encoding UTF8
$SoCthList = @()

$LastStartTime=0
forEach( $so in $($SrcObj | Select-String -pattern '([0-9]+)\=([0-9]+)\*([^\*]+)\*.*\r\n' -AllMatches).Matches ) {
    $SoCthList += $($so | Select-Object @{Label='ChapterUID';Expression={$so.groups[1].Value}},@{Label='ChapterTimeStart';Expression={$so.groups[2].Value}},@{Label='ChapString';Expression={$so.groups[3].Value}})
}
$SoCthList | Format-Table

#Converting to ogm formate and output to file.
$ogmList = ""
forEach( $so in $SoCthList) {
    $ogmList += $('CHAPTER{0:d2}' -f $so.ChapterUID) + '=' + $(Get-Date([long]$($so.ChapterTimeStart + '0000'))).ToString('HH:mm:ss.fff') + "`r`n"
    $ogmList += $('CHAPTER{0:d2}NAME' -f $so.ChapterUID) + '=' + $so.ChapString + "`r`n"
}
Write-Host $ogmList

$ogmList | Out-File -FilePath $OutputPath -Encoding utf8
