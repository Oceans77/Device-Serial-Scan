Remove-Job *

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "txt (*.txt)| *.txt"
    if($OpenFileDialog.ShowDialog() -eq "CANCEL"){
        break
    }
    return (Get-Content $OpenFileDialog.filename).Split()
}
$credential = Get-Credential
$list_of_computers = Get-FileName
$id_array = @()
New-Item $home\Desktop\Serial_list.csv -ItemType File -Force
Set-Content $home\Desktop\Serial_list.csv -Value "Computer Name, Serial"
$ErrorActionPreference = "SilentlyContinue"

$i = 0
foreach ($comp in $list_of_computers){
    Clear-Host
    $i++
    Write-Host "$i / $($list_of_computers.Count) -> $comp"
    Write-Host "Step 1: Starting Job"
    try{
        $id = Invoke-Command -ComputerName $comp -Credential $credential -ThrottleLimit 50 -AsJob -ScriptBlock {(Get-WmiObject win32_bios).Serialnumber}
        echo $id
    }catch{
        Continue
    }
    $id_array += $id.Id
}

$x = 0
foreach($job in $id_array){
    Clear-Host
    Write-Host "$x / $($id_array.Count) --> $($list_of_computers[$x])"
    Write-Host "Step 2: Writing Job Results to file"
    try{
        Wait-Job $job -Timeout 30 -ErrorAction Stop
    }catch{
        Add-Content $home\Desktop\Serial_list.csv -Value "$($list_of_computers[$x]),Offline"
        $x++
        continue
    }
    $serial = Receive-Job -Id $job
    if($serial -eq $null){
        Add-Content $home\Desktop\Serial_list.csv -Value "$($list_of_computers[$x]),Offline"
        $x++
        continue
    }
    Add-Content $home\Desktop\Serial_list.csv -Value "$($list_of_computers[$x]),$serial"
    $x++
}

Write-Host "COMPLETED" -ForegroundColor Green
