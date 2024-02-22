#coder: Andy Liao

#psm跟ps1同目錁E�E�E�E��E�E�E�好


$catchPsm =  $PSScriptRoot + "\webdrivercheck.psm1"

Import-Module $catchPsm 


$outputcsvpath = "\\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\webdriver_update_Tool\updatelog.csv"

#解壓縮路征E
$browsertypes=@("chrome","edge","firefox")
$brwDestinationPath ="\\192.168.60.16\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\selenium"

<##
$chromeDestinationPath = "\\192.168.60.16\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\selenium\chrome"
$edgeDestinationPath = "\\192.168.60.16\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\selenium\edge"
$firefoxDestinationPath = "\\192.168.60.16\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\selenium\firefox"
##>

$latestVersions = ""
$logs= $null

if (!(Test-Path $outputcsvpath)) {
    # 檔案不存在�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�進行後續處琁E
    $Title = "Type,Version,DateTime"
    # 中斁E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�加 -Encoding UTF8 | ft
    $Title | Add-Content -Path $outputcsvpath
    
foreach($browsertype in $browsertypes){

    #下載ZIP路征E
    echo "check&download $browsertype driver"

    $logs += download -browser $browsertype -outputpath $PSScriptRoot

}

}else{


    $data = Import-Csv -Path $outputcsvpath
    $logs = $null

foreach($browsertype in $browsertypes){

$lastdriverVersion =  ($data|?{$_.type -match $browsertype -and $_.Version.length -ne 0}|Sort-Object { [datetime]::ParseExact($_.DateTime, "MM/dd/yyyy HH:mm:ss", $null) }|select -Last 1).version

    <#$latestChromeVersions = $data | Where-Object {$_.Type -eq "Chromedriver"}

    # 篩選出 Type 為 Chromedriver, Edgedriver 戁EFirefoxdriver 且最新 version 皁E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�
    #$latestVersions = $data | 
    #Where-Object { $_.Type -in @("Chromedriver", "Edgedriver", "Firefoxdriver") } |
    #Group-Object Type | ForEach-Object {
    #    $_.Group | Select-Object -First 1
        #| Sort-Object {[version]::new($_.Version)} -Descending 
    # }

    $chromedriverVersion = $latestVersions | Where-Object Type -eq "Chromedriver" | Select-Object -ExpandProperty Version
    $edgedriverVersion = $latestVersions | Where-Object Type -eq "edgedriver" | Select-Object -ExpandProperty Version
    $firefoxdriverVersion = $latestVersions | Where-Object Type -eq "firefoxdriver" | Select-Object -ExpandProperty Version
    ##>

    #下載ZIP路征E
    echo "check&download $browsertype driver"

    $logs += download -browser $browsertype -outputpath $PSScriptRoot -lastVersion "$lastdriverVersion"
    $logs
}


}

if($logs.Length -ne 0){
 $logs.trim() | Add-Content -Path $outputcsvpath
 }

<# 建立一個陣列侁E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�存這三個字串
$array = @()
if($log -ne 0){
    $array += $log
}else{
    echo "chromedriver已經是最新版本"
}
if($log2 -ne 0){
    $array += $log2  
}else{
    echo "edgedriver已經是最新版本"
}
if($log3 -ne 0){
    $array += $log3
}else{
    echo "firefoxdriver已經是最新版本"
}

#$array = @($log, $log2, $log3)



if($array -ne $null){
    # 封E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�列匯出戁ECSV 檔桁E    echo "輸出log :" $outputcsvpath
    $array | Add-Content -Path $outputcsvpath
}


# 讀取現有的 CSV 檔桁E
#>


$dataSorted = Import-Csv -Path $outputcsvpath

# 依據 Version 欁E�E�E�E�E�E�E�E��E�E�E�E�E�E�E�排庁E
#$dataSorted = $dataSorted | Sort-Object DateTime -Descending
$dataSorted = $dataSorted | Sort-Object { [datetime]::ParseExact($_.DateTime, "MM/dd/yyyy HH:mm:ss", $null) } -Descending


#回存�E同一檔桁E
$dataSorted | Export-Csv -Path $outputcsvpath -NoTypeInformation -Encoding UTF8


#解壓ZIP路征E


$data = Import-Csv -Path $outputcsvpath


foreach($browsertype in $browsertypes){
 
  if($browsertype -match "firefox"){$zipname="gecko"}   ##special firefox driver name##
  else{$zipname=$browsertype} 

 $driverzip=(gci $PSScriptRoot\*.zip -Filter "*$zipname*" -ErrorAction SilentlyContinue).FullName

if($driverzip.Length -gt 0){
     
   $driverVersion = ($data|?{$_.type -match $browsertype}|Sort-Object { [datetime]::ParseExact($_.DateTime, "MM/dd/yyyy HH:mm:ss", $null) }|select -Last 1).version
   $drverfolder="$brwDestinationPath\$browsertype\$driverVersion\" 
   $drverfolder
    if(!(test-path $drverfolder)){New-Item -ItemType directory $drverfolder}
     
    Expand-Archive -Path $driverzip -DestinationPath $drverfolder -Force
    Remove-Item  $driverzip

    if(gci "$drverfolder\*"){
        $exepath = Get-ChildItem -Path $drverfolder -Recurse |Where-Object{$_.name -match "chromedriver.exe"} 
        $exepathfull=$exepath.FullName
        $exepathfull
        Move-Item -Path $exepathfull -Destination $drverfolder -Force
    }
        gci "$drverfolder\*" |?{$_.name -notmatch ".exe"}| Remove-item -Recurse -Force
}

#if folder than 3s, move file
if((Get-ChildItem -Path $brwDestinationPath\edge -Directory).Count -gt 3){
    $folder = Get-ChildItem -Path $brwDestinationPath\edge -Directory
    for($i=0;$i -lt $folder.Length-3;$i++){
        $movepath = $brwDestinationPath + "\edge\" + $folder[$i].Name
        Move-Item -Path $movepath -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\webdriver_update_Tool\olddriver\edge"    
    }
}
if((Get-ChildItem -Path $brwDestinationPath\firefox -Directory).Count -gt 3){
    $folder = Get-ChildItem -Path $brwDestinationPath\firefox -Directory
    for($i=0;$i -lt $folder.Length-3;$i++){
        $movepath = $brwDestinationPath + "\firefox\" + $folder[$i].Name
        Move-Item -Path $movepath -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\webdriver_update_Tool\olddriver\firefox"      
    }
}
if((Get-ChildItem -Path $brwDestinationPath\chrome -Directory).Count -gt 3){
    $folder = Get-ChildItem -Path $brwDestinationPath\chrome -Directory
    for($i=0;$i -lt $folder.Length-3;$i++){
        $movepath = $brwDestinationPath + "\chrome\" + $folder[$i].Name
        Move-Item -Path $movepath -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\webdriver_update_Tool\olddriver\chrome"      
    }
}


}

$des20="\\192.168.20.20\sto\EO\2_AutoTool\ALL\103.Dell_AITest\selenium"
$sizefrom= (Get-ChildItem $brwDestinationPath -recurse -force  | Measure-Object -property Length -sum | Select-Object Sum).sum
$sizeto= (Get-ChildItem  $des20 -recurse -force  | Measure-Object -property Length -sum | Select-Object Sum).sum
if($sizefrom -ne $sizeto ){
remove-item $des20 -Recurse -Force -ErrorAction SilentlyContinue
copy-item $brwDestinationPath -Destination (split-path $des20) -Recurse -Force -ErrorAction SilentlyContinue
}