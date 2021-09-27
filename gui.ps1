#We need to load this assembly to create a gui
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
#helper functions
Function Get-Folder($initialDirectory="$PSScriptRoot")

{

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select WAD folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
Function Write-YesNoPrompt{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Question,
        [Parameter(Mandatory = $true)]
        [String]
        $Title

    )

    $wshell = New-Object -ComObject Wscript.Shell
    $answer = $wshell.Popup($question,0,$title,64+4)

    if($answer -eq 6){
        return 'y'
    }
    else{
        return 'n'
    }
    
}

Function Write-InformationalPrompt{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Information,
        [Parameter(Mandatory = $true)]
        [String]
        $Title
    )

    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($Information,0,$Title,64+0)

    return
}
Function Write-AlertPrompt{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Alert
    )

    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($Alert,0,"Error",16+0)

    return
}
function Write-LauncherConfig{

    Write-InformationalPrompt -Information "Please select the folder for you WADs" -Title "Select WAD folder"

    if($a -eq 'n'){
        Exit
    }
    $WADPath = Get-Folder

    #Generate config file by grabbing all file names in the folder specified. Then checking for the .WAD extension. Name the launch option the name of the file
    $filesInWadFolder = (Get-ChildItem -Path $WADPath).Name

    $filesToAddToCfg = @()
    foreach($file in $filesInWadFolder){
        #Filename only matters if we have the .WAD extension so we don't really care if we cut it short for other file extensions
        $ext = $file[($file.Length-4)..($file.Length-1)] -join ''
        $filename = $file[0..($file.Length-5)] -join ''

        if($ext -eq ".WAD"){
            $temphash = @{"filename" = $filename; "fullpath" = "$WADPath\$file"}
            $filesToAddToCfg+=$temphash
        }
    }
    #Make sure we're not just creating an empty config file:
    if($filesToAddToCfg.Length -eq 0){
        Write-AlertPrompt -Alert "No .WAD files found in specified folder. Please check the folder and try again"
        Exit
    }

    #Create the config file and add the information to it
    New-Item -Path "$PSScriptRoot\launcher.cfg" | Out-Null

    foreach($item in $filesToAddToCfg){
        $name = $item["filename"]
        $path = $item["fullpath"]
        Add-Content -Path "$PSScriptRoot\launcher.cfg" -Value "$name`;$path"
    }

    Write-InformationalPrompt -Information "Config file should have been generated succesfully at $PSScriptRoot\launcher.cfg" -Title "Success!"
}

#main loop
$configExists = Test-Path "$PSScriptRoot\launcher.cfg"
if(!$configExists){
    $a = Write-YesNoPrompt -Question "No config file found, would you like to create a new one?" -Title "No config file found"
    Write-LauncherConfig
}

#Check the current config file and if there are updates to the folder:
$launcherContent = Get-Content "$PSScriptRoot\launcher.cfg"

$currentWads = @()
foreach($string in $launcherContent){
    $currentWads+= @{"filename" = $string.split(";")[0]; "fullpath" = $string.split(";")[1]}
}
#Grab a fullpath to find the WAD folder
$WADFolder = $currentWads[0]["fullpath"][0..($currentWads[0]["fullpath"].LastIndexOf('\')-1)] -join ''

#Iterate through the WAD folder to find any names that might be there which have been missed or added:
#NOTE: I tried to use the contains keyword for this but for some reason it just doesn't work with the dictionary object properly. Which explains the extra foreach loop.

$filesInWadFolder = (Get-ChildItem -Path $WADFolder).Name
$newWads = @()
foreach($file in $filesInWadFolder){
    $ext = $file[($file.Length-4)..($file.Length-1)] -join ''
    $filename = $file[0..($file.Length-5)] -join ''
    $fullpath = "$WADFolder\$filename"
    if($ext -eq ".WAD"){
        $exists = $false
        foreach($item in $currentWads){
            if($filename -eq $item["filename"]){
                $exists = $true
                break
            }
        }
        if(!$exists){
            $newWads += @{"filename" = $filename; "fullpath" = $fullpath}
        }
    }
}

#Then add the new WADs to the config file and add them to our names and fullpaths lists
foreach($newWad in $newWads){
    $name = $newWad["filename"]
    $path = $newWad["fullpath"]
    Write-Host "Main functionality, writing $name;$path to cfg file"
    Add-Content -Path "$PSScriptRoot\launcher.cfg" -Value "$name`;$path"
    $currentWads+=@{"filename" = $name;"fullpath" = $path}
}

#Now we can spawn the gui

#Main window
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Crispy Doom WAD launcher'
$main_form.Width = 600
$main_form.Height = 400
$main_form.AutoSize = $true

#WADs label
$Label = New-Object System.Windows.Forms.Label
$Label.Location =  New-Object System.Drawing.Point(0,10)
$Label.Dock = "Fill"
$Label.Text = "WAD to run"
$Label.AutoSize = $true
$Label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,[System.Drawing.FontStyle]::Regular)
$main_form.Controls.Add($Label)

#Dropdown select of WADs
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 300
$ComboBox.Location = New-Object System.Drawing.Point(0,25)
$ComboBox.AutoSize = $true

Foreach ($WAD in $currentWads)
{
    $ComboBox.Items.Add($WAD["filename"]);
}
$main_form.Controls.Add($ComboBox)

#Make new launcher config button
$LauncherCfgButton = New-Object System.Windows.Forms.Button
$LauncherCfgButton.Size = New-Object System.Drawing.Size(120,23)
$LauncherCfgButton.Location = New-Object System.Drawing.Point(0,40)
$LauncherCfgButton.Text = "Create new launcher config"
$LauncherCfgButton.Dock = "Bottom"
$LauncherCfgButton.AutoSize = $true
$LauncherCfgButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",30,[System.Drawing.FontStyle]::Regular)

$LauncherCfgButton.Add_Click(
{
    $a = Write-YesNoPrompt -Title "Warning" -Question "This will delete your old launcher config file. Are you sure?"
    if($a -eq 'y'){
        Remove-Item -Path "$PSScriptRoot\launcher.cfg"
        Write-LauncherConfig
    }
}
)

$main_form.Controls.Add($LauncherCfgButton)


#Run button
$RunButton = New-Object System.Windows.Forms.Button
$RunButton.Size = New-Object System.Drawing.Size(120,23)
$RunButton.Text = "Run"
$RunButton.Dock = "Bottom"
$RunButton.AutoSize = $true
$RunButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",30,[System.Drawing.FontStyle]::Regular)
$RunButton.BackColor = "Green"
$RunButton.Add_Click(
{
    $selectediWad = $ComboBox.SelectedItem
    $iwadPath = ""
    foreach($item in $currentWads){
        if($item["filename"] -eq $selectediWad){
            $iwadPath = $item["fullpath"]
        }
    }
    if($iwadPath.Length -eq 0){
        Write-AlertPrompt -Alert "Please select a WAD"
    }
    else{
        Start-Process -FilePath "$PSScriptRoot\crispy-doom.exe" -ArgumentList "-iwad $iwadPath"
    }
}
)
$main_form.Controls.Add($RunButton)

$main_form.ShowDialog()