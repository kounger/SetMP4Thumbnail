#Edit this variable to set a directory as the the current working location:
#(E.g $working_location  = "C:\Users\User\MyVideos")
#This is optional but makes it possible to work without full file paths.
#($PSScriptRoot is the directory where this powershell script is located.)
$working_location = $PSScriptRoot

#EDIT these Variables:
$video_path = "foobar.mp4"           #Must be a file with .mp4 extension. 
$timestamp  = "00:01:42.500"         #Can be ...HH:MM:SS.XXXXXX... or ...S.XXXXXX... format.


####################################################################################################


Write-Output( "`n" + "#################" `
            + "`n" + "SET-MP4-THUMBNAIL" `
            + "`n" + "#################")


#Set the current working location:
if (Test-Path -Path $working_location -PathType Container)
{
    #Change the current location for this PowerShell runspace:
    Set-Location $working_location

    #Change the current working directory for .NET:
    [System.IO.Directory]::SetCurrentDirectory($working_location)
} 
else 
{
    throw [IO.FileNotFoundException] ("Specified working location is not a directory or does not exist:" `
                                     +"`n" + $working_location) 
}


#Check if FFmpeg is installed on this system:
try
{
    ffmpeg -version | Out-Null
}
catch [System.Management.Automation.CommandNotFoundException]
{
    throw [System.Management.Automation.CommandNotFoundException] `
          ("FFmpeg is not installed on this system!" `          +"`nInstall FFmpeg for this script to function!")
}


#Get the full path to the video file:
$video_path = [IO.Path]::GetFullPath($video_path)


#Check if video format is MP4:
#(Using this script on other formats will likely ruin the file.)
$format = [System.IO.Path]::GetExtension($video_path)

if (-not($format -match ".mp4"))
{
    throw [NotSupportedException] ("This script only works with MP4 files.")
}


#Check if the specified video file exists:
if (-not([System.IO.File]::Exists($video_path))) 
{
    throw [IO.FileNotFoundException] ("The specified file does not exist:" `
                                     +"`n" + $video_path) 
}


#Check if correct time format has been used:
$time_regex_one = '^\d{2,}:[0-5][0-9]:[0-5][0-9](\.\d+)?$'  #...HH:MM:SS.XXXXXX...
$time_regex_two = '^\d+(\.\d+)?$'                           #...S.XXXXXX...

if (-not(($timestamp -match $time_regex_one) -or ($timestamp -match $time_regex_two)))
{
    throw [FormatException] ("The specified timestamp has a wrong syntax:" `
                            +"`n" + $timestamp)
}


#Transform ...HH:MM:SS.XXXXXX... into ...S.XXXXXX... format:
if ($timestamp -match $time_regex_one)
{
    [Double[]]$split_timestamp = $timestamp.Split(':')
              $timestamp       = $split_timestamp[0] * 3600 + $split_timestamp[1] * 60 + $split_timestamp[2]
}


#Check if the specified time exceeds the video duration:
$video_duration = ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $video_path

if ($timestamp -gt $video_duration)
{
    throw [ArgumentOutOfRangeException] ("The specified time in `$timestamp is larger than the video duration:" `
                                        +"`n" + $timestamp + " seconds" + "  >  " + $video_duration + " seconds")        
} 


#Check if the target hard drive has enough space for this operation: 
$free_space = Get-Volume -FilePath $video_path | Select -ExpandProperty SizeRemaining
$video_size = Get-Childitem -File  $video_path | Select -ExpandProperty Length
$target_folder = Split-Path $video_path -Parent
$space_needed  = "{0:N2} MB" -f ($video_size / 1MB)

if ($video_size -gt $free_space)
{
    throw [System.IO.IOException] ("There is not enough disk space for this operation to continue:" `
                                  +"`n" + "$space_needed of free space needed in `"$target_folder`".")
}


#This function asks the user to close a certain file if it's open in the background:
function CheckForOpenFile{
    param(
    [Parameter(Mandatory)]
    [string]$File
    )    
    
    try 
    {
        [IO.File]::OpenWrite($File).close()
    }
    catch [System.IO.IOException] 
    {    
        $video_name = [IO.Path]::GetFileName($File)
        
        Write-Output ("`n")
        Write-Warning("`nThe operation can't continue because the video file `"$video_name`" is open in the background.")        
        
        $confirm = Read-Host "`nClose the video file and presss ENTER to continue"

        #Repeat until the file is closed:
        CheckForOpenFile -File $File
    }
}


#Throw an error if the Video-File is still open somewhere:
CheckForOpenFile -File $video_path


#Create a screenshot at a defined time into the video:
Write-Output("`n" * 2 + "CREATE THUMBNAIL:" + "`n")

$video_directory = [IO.Path]::GetDirectoryName($video_path)
$thumb_name      = "temp_screen_" + [guid]::NewGuid().ToString() + ".jpg"
$temp_thumb_file = Join-Path $video_directory -ChildPath $thumb_name
ffmpeg -ss $timestamp -i $video_path -vframes 1 -q:v 2 -y $temp_thumb_file 2>&1 | % {"$_"}


#Rename the old file while it's pending for deletion: 
$file_name = [IO.Path]::GetFileNameWithoutExtension($video_path)
$extension = [IO.Path]::GetExtension($video_path)
$old_video_name = "old_" + $file_name + $extension + "_" + [guid]::NewGuid().ToString() + $extension
$old_video_path = Join-Path $video_directory -ChildPath $old_video_name

Rename-Item -Path $video_path -NewName $old_video_name


#Set the screenshot as the new video thumbnail:
Write-Output("`n" * 2 + "ENCODE THUMBNAIL INTO THE VIDEO:" + "`n")

ffmpeg -i $old_video_path -i $temp_thumb_file -map 1 -map 0 -c copy -disposition:0 attached_pic -y $video_path 2>&1 | % {"$_"}


#Present the new thumbnail to the user by opening up the folder of the file:

#First close all open explorer windows of this folder:
#(Otherwise the folder windows will start stacking up while using this script.)
$shell  = New-Object -ComObject Shell.Application
$window = $shell.Windows() | Where-Object {$_.Document.Folder.Self.Path -eq $video_directory}
$window | ForEach-Object {$_.Quit()}

Start-Process -FilePath C:\Windows\explorer.exe -ArgumentList "/select, ""$video_path"""


#Ask the user if he wants to keep the new thumbnail:
$user_input = Read-Host "`n THUMBNAIL UPDATED:" `
                        "`n Enter 'y' to keep the updated file." `
                        "`n Enter 'n' to revert the changes." `
                        "`n Your Option (y/n)"

while (-not($user_input -eq 'y' -or $user_input -eq 'n'))
{
    $user_input = Read-Host "`n Only 'y' or 'n' are permitted answers:" `
                            "`n Your Option (y/n)"
}


#Make sure the updated video file and the the original file are closed:
CheckForOpenFile -File $video_path
CheckForOpenFile -File $old_video_path


#Keep the updated file and delete the old file:
if ($user_input -eq 'y')
{
    #Delete the old video file:
    #Remove-Item $old_video_path

    #Send the old video file to the recycle bin:
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($old_video_path,'OnlyErrorDialogs','SendToRecycleBin')

    Remove-Item $temp_thumb_file

    Write-Host "`nThe file has received a new thumbnail. The original file has been moved to the Recycle Bin." `
                -ForegroundColor Green
}
#Remove the updated file and reinstate the old file:
elseif ($user_input -eq 'n')
{
    #Remove the updated file:
    Remove-Item $video_path   

    #Give the original file it's old name back:
    Rename-Item -Path $old_video_path -NewName ([IO.Path]::GetFileName($video_path))

    Remove-Item $temp_thumb_file

    Write-Host "`nThe original file has been restored. The updated file has been deleted." `
                -ForegroundColor Green
}