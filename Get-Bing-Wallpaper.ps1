###########################################
#
# 本PowerShell脚本参考了以下两位作者的文章并且融合。并加以汉化而成
# https://blog.hyperexpert.com/powershell-script-to-download-bings-daily-image-as-a-wallpaper/  (Ali Khalidy)
# 以及
# http://www.pstips.net/set-wallpaper-by-bing-image.html  (Mooser Lee)
#
# 站在巨人肩上改脚本，特此感谢！ Thanks for the authors above!
#
###########################################
$Market = "zh-CN"
$Resolution = "1920x1080"
$DownloadDirectory = "$env:USERPROFILE\Pictures\Bing Wallpaper"
$bingImageApi = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=$($Market))"

While (!(Test-Connection -ComputerName bing.com -count 1 -Quiet -ErrorAction SilentlyContinue )) {
    Write-Host -ForegroundColor Red "等待网络连接..."
    Start-Sleep -Seconds 10
}

New-Item -ItemType directory -Force -Path $DownloadDirectory | Out-Null

[xml]$Bingxml = (Invoke-WebRequest -Uri $bingImageApi).Content
$ImageUrl = "http://www.bing.com$($Bingxml.images.image.urlBase)_$($Resolution).jpg";

$ImageFileName = "$($Bingxml.images.image.fullstartdate).jpg"
$BingImageFullPath = "$($DownloadDirectory)\$($ImageFileName)"

if ((Test-Path "$BingImageFullPath") -And ((Get-ChildItem "$BingImageFullPath").LastWriteTime.ToShortDateString() -eq (get-date).ToShortDatesTring())){
    Write-Host -ForegroundColor Green "今天已经下载了最新的必应首页背景图片，存放路径 $DownloadDirectory"   
}
else {
    Invoke-WebRequest -UseBasicParsing -Uri $ImageUrl -OutFile "$BingImageFullPath";
    Write-Host -ForegroundColor Green "必应首页背景图片已经下载到 $DownloadDirectory" 
}

While (!(Test-Path "$BingImageFullPath")) {
    Write-Host -ForegroundColor Yellow "等待图片下载中..."
    Start-Sleep -Seconds 10
}
Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
   public enum Style : int
   {
       Tile, Center, Stretch, NoChange
   }
   public class Setter {
      public const int SetDesktopWallpaper = 20;
      public const int UpdateIniFile = 0x01;
      public const int SendWinIniChange = 0x02;
      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
      private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
      public static void SetWallpaper ( string path, Wallpaper.Style style ) {
         SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
         RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
         switch( style )
         {
            case Style.Stretch :
               key.SetValue(@"WallpaperStyle", "2") ; 
               key.SetValue(@"TileWallpaper", "0") ;
               break;
            case Style.Center :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "0") ; 
               break;
            case Style.Tile :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "1") ;
               break;
            case Style.NoChange :
               break;
         }
         key.Close();
      }
   }
}
"@
[Wallpaper.Setter]::SetWallpaper( "$BingImageFullPath", 3 )