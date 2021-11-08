# This function requests the run path in a folder tree gui

$DesktopPath = [Environment]::GetFolderPath("Desktop")

Function Get-Folder($initialDirectory="$DesktopPath\<<folder>>"){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select the correct Display folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory
    $foldername.ShowNewFolderButton = $true
    $result = $foldername.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
    if ($result -eq [Windows.Forms.DialogResult]::OK){
        $folder = $foldername.SelectedPath
    }
    return $folder
}

# setting up variables to grab the correct path to the data, call an empty array for the list of leaf directories, and add the compression .NET class to the powershell session
$sourceParent = Get-Folder
$source = Get-ChildItem -Path $sourceParent -Directory -Recurse
$leafdirs = New-Object System.Collections.ArrayList
Add-Type -assembly "system.io.compression.filesystem"

# This loop grabs all the leaf directories in the selected path and assigns them into the $leafdirs array
Foreach ($s in $source) {
  if ((Get-ChildItem $s.fullname -Directory).count -eq 0) {
    $leafdirs.add($s.fullname)
  }
}

# This loop grabs the last 4 directories in the path and uses them to name and create the zipfile in the ZipOutput directory on the Desktop
Foreach ($l in $leafdirs) {
  $zipfile = ((($l | Split-Path -NoQualifier).split('\') | Select-Object -last 4) -join ' ')
  $destination = Join-path -path $DesktopPath\ZipOutput\ -ChildPath "$zipfile.zip"
  # If(Test-path $destination) {Remove-item $destination}
  [io.compression.zipfile]::CreateFromDirectory($l, $destination)
}
