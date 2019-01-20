$host.ui.RawUI.WindowTitle = "Alup Batch Converter"
Write-Host Please enter the path to the root directory where the files reside.
$path = Read-Host "Enter the path"
$filelist = Get-ChildItem $path -Exclude '*.mp4','Season*','Special*' -recurse
 
$num = $filelist | measure
$filecount = $num.count
 
$i = 0;
ForEach ($file in $filelist)
{
    $i++;
    $oldfile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;
    $newfile = $file.DirectoryName + "\" + $file.BaseName + ".mp4";
      
    $progress = ($i / $filecount) * 100
    $progress = [Math]::Round($progress,2)
 
    Clear-Host
    Write-Host -------------------------------------------------------------------------------
    Write-Host Alup Batch Converter
    Write-Host "Processing - $oldfile" 
    Write-Host "File $i of $filecount - $progress%"
    Write-Host -------------------------------------------------------------------------------
  if (Test-Path HandBrakeCLI.exe)
  { 
    Start-Process "HandBrakeCLI.exe" -ArgumentList "-i `"$oldfile`" -o `"$newfile`" -f mp4  -O -e x264 -q 24 -E aac -6 stereo -R 44.1 -B 96k-x cabac=1:ref=5:analyse=0x133:me=umh:subme=9:chroma-me=1:deadzone-inter=21:deadzone-intra=11:b-adapt=2:rc-lookahead=60:vbv-maxrate=10000:vbv-bufsize=10000:qpmax=69:bframes=5:b-adapt=2:direct=auto:crf-max=51:weightp=2:merange=24:chroma-qp-offset=-1:sync-lookahead=2:psy-rd=1.00,0.15:trellis=2:min-keyint=23:partitions=all" -Wait -NoNewWindow
    Wait-Process -Name HandBrakeCLI.exe
    #Remove-Item -Force $oldfile
  }
  else
  {
    Clear-Host
    write-host Error HandBrakeCLI.exe not found! $objItem.Name, $objItem.WorkingSetSize -foregroundcolor "red"
    Write-Host Press any key to exit...
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Exit
  }
}