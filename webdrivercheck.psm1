#coder: Andy Liao

function edgelatestversion{

  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
    $wshell=New-Object -ComObject wscript.shell
      Add-Type -AssemblyName Microsoft.VisualBasic
      Add-Type -AssemblyName System.Windows.Forms

 function Set-WindowState {
	<#
	.LINK
	https://gist.github.com/Nora-Ballard/11240204
	#>

	[CmdletBinding(DefaultParameterSetName = 'InputObject')]
	param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[Object[]] $InputObject,

		[Parameter(Position = 1)]
		[ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE',
					 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED',
					 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
		[string] $State = 'SHOW'
	)

	Begin {
		$WindowStates = @{
			'FORCEMINIMIZE'		= 11
			'HIDE'				= 0
			'MAXIMIZE'			= 3
			'MINIMIZE'			= 6
			'RESTORE'			= 9
			'SHOW'				= 5
			'SHOWDEFAULT'		= 10
			'SHOWMAXIMIZED'		= 3
			'SHOWMINIMIZED'		= 2
			'SHOWMINNOACTIVE'	= 7
			'SHOWNA'			= 8
			'SHOWNOACTIVATE'	= 4
			'SHOWNORMAL'		= 1
		}

		$Win32ShowWindowAsync = Add-Type -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

		if (!$global:MainWindowHandles) {
			$global:MainWindowHandles = @{ }
		}
	}

	Process {
		foreach ($process in $InputObject) {
			if ($process.MainWindowHandle -eq 0) {
				if ($global:MainWindowHandles.ContainsKey($process.Id)) {
					$handle = $global:MainWindowHandles[$process.Id]
				} else {
					Write-Error "Main Window handle is '0'"
					continue
				}
			} else {
				$handle = $process.MainWindowHandle
				$global:MainWindowHandles[$process.Id] = $handle
			}

			$Win32ShowWindowAsync::ShowWindowAsync($handle, $WindowStates[$State]) | Out-Null
			Write-Verbose ("Set Window State '{1} on '{0}'" -f $MainWindowHandle, $State)
		}
	}
}

$cSource = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class Clicker
{
//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct INPUT
{ 
    public int        type; // 0 = INPUT_MOUSE,
                            // 1 = INPUT_KEYBOARD
                            // 2 = INPUT_HARDWARE
    public MOUSEINPUT mi;
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT
{
    public int    dx ;
    public int    dy ;
    public int    mouseData ;
    public int    dwFlags;
    public int    time;
    public IntPtr dwExtraInfo;
}

//This covers most use cases although complex mice may have additional buttons
//There are additional constants you can use for those cases, see the msdn page
const int MOUSEEVENTF_MOVED      = 0x0001 ;
const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
const int MOUSEEVENTF_LEFTUP     = 0x0004 ;
const int MOUSEEVENTF_RIGHTDOWN  = 0x0008 ;
const int MOUSEEVENTF_RIGHTUP    = 0x0010 ;
const int MOUSEEVENTF_MIDDLEDOWN = 0x0020 ;
const int MOUSEEVENTF_MIDDLEUP   = 0x0040 ;
const int MOUSEEVENTF_WHEEL      = 0x0080 ;
const int MOUSEEVENTF_XDOWN      = 0x0100 ;
const int MOUSEEVENTF_XUP        = 0x0200 ;
const int MOUSEEVENTF_ABSOLUTE   = 0x8000 ;

const int screen_length = 0x10000 ;

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
[System.Runtime.InteropServices.DllImport("user32.dll")]
extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

public static void LeftClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}
}
'@
Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing


### start to setting##
function openedge{
Start-Process msedge.exe 
 start-sleep -s 20
 $id=(Get-Process msedge |?{($_.MainWindowTitle).length -gt 0}).Id
  start-sleep -s 2
  Get-Process -id $id | Set-WindowState -State MAXIMIZE

[Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
 start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")
 start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")
  start-sleep -s 2
    
  do{
     [Clicker]::LeftClickAtPoint(100,100)
  start-sleep -s 2
   Set-Clipboard -Value "edge://settings/help"
   start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait("^l")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^v")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
  start-sleep -s 5
 [Clicker]::LeftClickAtPoint(100,100)
  start-sleep -s 2
  [System.Windows.Forms.SendKeys]::SendWait("%e")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("b")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("m")
  start-sleep -s 5
  [System.Windows.Forms.SendKeys]::SendWait("^l")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^a")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^c")
  start-sleep -s 2
  $webadd2=Get-Clipboard
   start-sleep -s 2
   }until($webadd2 -like "edge://settings/help")

  }

  openedge

    $reconnect_time=0

  do{   

    $id=(Get-Process msedge |?{($_.MainWindowTitle).length -gt 0}).Id
    if(!$id){openedge}
     $id=(Get-Process msedge |?{($_.MainWindowTitle).length -gt 0}).Id
   [Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
   start-sleep -s 2

   Set-Clipboard -Value "about"
   start-sleep -s 5
    
  [System.Windows.Forms.SendKeys]::SendWait("^f")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^v")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")

 start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^a")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^c")
  start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait("{esc}")

$content=Get-Clipboard
  start-sleep -s 5
 # $content

 if($content -like "* then refresh the page*"){
 [System.Windows.Forms.SendKeys]::SendWait("{F5}")
   start-sleep -s 30
 }

 if($content -like "*Unable to connect to the Internet*"){
 ipconfig /renew
  start-sleep -s 10
  $reconnect_time++  

[System.Windows.Forms.SendKeys]::SendWait("{f5}")
 
   Set-Clipboard -Value "about"
   start-sleep -s 5
    
[Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
   start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^f")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^v")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")

 start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^a")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^c")
  start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait("{esc}")

$content=Get-Clipboard
  start-sleep -s 5


 }
 if($content -like "*restart Microsoft Edge*") {   

   Set-Clipboard -Value "restart"
   start-sleep -s 5
    
[Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
   start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^f")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^v")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
 start-sleep -s 10

}



}until($content -like "*restart Microsoft Edge*" -or $content -like "*Microsoft Edge is up to date*"  -or $reconnect_time -gt 1 -or $content -like "*Unable to connect to the internet*")

if(!($content -like "*Microsoft Edge is up to date*" -and $reconnect_time -le 1) -and !($content -like "*Unable to connect to the internet*") ){
 



 ###### check version ###
  
    $id=(Get-Process msedge |?{($_.MainWindowTitle).length -gt 0}).Id
    if(!$id){openedge}
     $id=(Get-Process msedge |?{($_.MainWindowTitle).length -gt 0}).Id

   Set-Clipboard -Value "about"
   start-sleep -s 5
    
[Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
   start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^f")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^v")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("~")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("{esc}")
 start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^a")
  start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait("^c")
  start-sleep -s 5
$content=Get-Clipboard

}

if($content -like "*Microsoft Edge is up to date*"  -and $reconnect_time -le 1 ){

$indexline=$content -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}"

$index=($indexline.split(" ")) -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}"

$index

}

#return $index

 (get-process -name msedge).CloseMainWindow() |out-null

}


    function download([string]$browser,[string]$outputpath,[string]$lastVersion){



        if($browser -eq "chrome"){
    
             
            $url = "https://googlechromelabs.github.io/chrome-for-testing/#stable"
            
            $response = Invoke-WebRequest -Uri $url
            
           
            #check version
            $allelement = $response.AllElements
            foreach($items in $allelement){
                if($items.class -eq "status-ok" -and $items.innerText -match "Stable"){
                    $pattern = '\d+\.\d+\.\d+\.\d+'
                    $match = [regex]::Match($items.outerText, $pattern)
                }
            }
            $version = $match.Value


            #catch downloadlist
            $text = $response.Content
            $regex = 'https://.*?.zip'
            $links = Select-String -InputObject $text -Pattern $regex -AllMatches | Foreach-Object {$_.Matches} | Foreach-Object {$_.Value}
            $fliter = @()

            foreach ($link in $links) {
                $fliter += $link
            }

            $fliter = $fliter | Select-String -Pattern $version

            
            $downloadurl = (($fliter | Select-String -Pattern "chromedriver-win64" |Get-Unique)|out-string).trim()
            $uri = New-Object -TypeName System.Uri -ArgumentList $downloadurl

            if($lastVersion -ne $version){

                $outputPath = $outputPath + "\chromedriver.zip"
                Invoke-WebRequest -Uri $uri -OutFile $outputPath

                $now = Get-Date
                $log = "Chromedriver,$version,$now `n"
                return $log
            }#else{
                #return $null
            #}


            
        }

        if($browser -eq "edge"){

        
             $version=edgelatestversion
             
            # https://msedgedriver.azureedge.net/$version/edgedriver_win64.zip
            #$url = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
           
            #$response = Invoke-WebRequest -Uri $url
            #$text = $response.Content

           

            #$regex = '<a\s+.*?>.*?</a>'
            #$links = Select-String -InputObject $text -Pattern $regex -AllMatches | Foreach-Object {$_.Matches} | Foreach-Object {$_.Value}
            #$fliter = @()

            #foreach ($link in $links) {
            #   if($link -match "win64.zip"){
            #       $fliter += $link
            #   }
            #}


            # 使用正規表達式替換除亁Ehref 屬性中皁E�E��E�串以外的內容
            #$downloadurl = $fliter[0] -replace '.*href="([^"]+)".*', '$1'
            
            #$outputPath = $outputPath + "\edgedriver.zip"

            #$version = $downloadurl -split "/"
            #$version = $version[3]
            $url = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
           
            $response = Invoke-WebRequest -Uri $url
            $text = $response.Content

            $tmp=""
           # 查找字符串中是否匁E�E��E�特定的子字符串
            if ($text.IndexOf('<a href="https://msedgedriver.azureedge.net/') -ge 0) {
                foreach($sp in $text[($text.IndexOf('<a href="https://msedgedriver.azureedge.net/')+9)..($text.IndexOf('<a href="https://msedgedriver.azureedge.net/')+77)]){
                    $tmp += $sp
                }
            } 
            $tmp = $tmp -split "/"

            #$version = $tmp[3]


            $downloadurl = "https://msedgedriver.azureedge.net/" +  $version + "/edgedriver_win64.zip"
            
            $outputPath = $outputPath + "\edgedriver.zip"
             

            if($lastVersion -ne $version){
                Invoke-WebRequest -Uri $downloadurl -OutFile $outputPath
                $now = Get-Date
                $log = "edgedriver," + $version + "," + $now+"`n"

                return $log
            }#else{
                #return $null
            #}



            
        }

         if($browser -eq "firefox"){
            #$url = "https://github.com/mozilla/geckodriver/releases"
            #$outputPath = "C:\example.txt"
            #$response = Invoke-WebRequest -Uri $url
            #$text = $response.Content

            #$regex = '<a\s+.*?\bhref="https:\/\/msedgedriver\.azureedge\.net\</a>'

            #$regex = 'href=".*?/mozilla/geckodriver/releases/download/v([\d\.]+)/geckodriver-v[\d\.]+-win-aarch64.zip"'
            #$links = Select-String -InputObject $text -Pattern $regex -AllMatches | Foreach-Object {$_.Matches} | Foreach-Object {$_.Value}

            #$versions = @()
            #foreach ($link in $links) {
            #    $version = $link -replace '.*v([\d\.]+).*', '$1'
            #}
            $url = "https://github.com/mozilla/geckodriver/releases"
            #$outputPath = "C:\example.txt"
            $response = Invoke-WebRequest -Uri $url
            $text = $response.Content

            $regex = '<a\s+.*?>.*?</a>'
            $links = Select-String -InputObject $text -Pattern $regex -AllMatches | Foreach-Object {$_.Matches} | Foreach-Object {$_.Value}
            $fliter = @()

            foreach ($link in $links) {
               if($link -match "/tag/v"){
                   $fliter += $link
               }
            }
            $version = $fliter[0] -replace '.*v([\d\.]+).*', '$1'


            if($lastVersion -ne $version){
                # 使用正規表達式替換除亁Ehref 屬性中皁E�E��E�串以外的內容
               # $downloadurl = $links -replace '.*href="([^"]+)".*', '$1'
                #$downloadurl = "https://github.com" + $downloadurl
                $downloadurl = "https://github.com/mozilla/geckodriver/releases/download/v"+ $version + "/geckodriver-v" + $version + "-win64.zip"
                $outputPath = $outputPath + "\geckodriver.zip"
                

                Invoke-WebRequest -Uri $downloadurl -OutFile $outputPath

                $now = Get-Date
                $log = $version + "," + $now

                return "firefoxdriver," + $log+"`n"
            }#else{
               # return $null
            #}
            
         }
    }


# 匯出模絁E�E�E員
Export-ModuleMember -Function download




