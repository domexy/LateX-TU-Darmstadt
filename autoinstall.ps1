#----------------------------
# Link and Name Declarations 
#----------------------------

$link_design = "https://www.tu-darmstadt.de/media/medien_stabsstelle_km/services/medien_cd/latex/latex-tudesign_2016-03-01.zip"
$name_design = "tu-design"
$link_thesis = "https://www.tu-darmstadt.de/media/medien_stabsstelle_km/services/medien_cd/latex/latex-tudesign-thesis_0020140703.zip"
$name_thesis = "tu-thesis"
$link_template = "https://www.tu-darmstadt.de/media/medien_stabsstelle_km/services/medien_cd/latex/Vorlage_Dissertation_2016-03-01.zip"
$name_template = "template"
$link_fonts = "https://www.tu-darmstadt.de/media/medien_stabsstelle_km/services/medien_cd/tu-darmstadt-schriften_jan08.zip"
$name_fonts = "fonts"

$link_miktex_x64 = "http://mirror.ctan.org/systems/win32/miktex/setup/windows-x64/miktexsetup-x64.zip"
$link_miktex_x86 = "http://mirror.ctan.org/systems/win32/miktex/setup/windows-x86/miktexsetup.zip"
$name_miktex = "miktex-setup"

#----------------------------
# Function Declarations 
#----------------------------

function download{
    param( $url, $dl_name )
    Write-Host "------------";
    Write-Host "Downloading $dl_name";
    Write-Host "From $url";

    Invoke-WebRequest -Uri $url -OutFile "$dl_name.zip"
    if ((Test-Path "$dl_name.zip") -eq $True){
        
		$file_path=(Get-Item -Path ".\" -Verbose).FullName + "\$dl_name.zip"
        Write-Host "to $file_path";
	}else{
		write-host "ERROR: $dl_name could not be downloaded" -foreground "red"
    }
   
}

function Is-Installed( $program ) {
    
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;

    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;

    return $x86 -or $x64;
}

#----------------------------
# Main Program 
#----------------------------

#---Define scope of program---
# Installing classes only or necessary programs also

Write-Host "
[1] (Re-)Install MikTex, TeXWorks and TU-Darmstadt classes
[2] Install MikTex, TeXWorks (if missing) and TU-Darmstadt classes
[3] Only install TU-Darmstadt classes"  -foreground "yellow"
$valid_selection = $False;
while($valid_selection -ne $True)
{
    $inst_selection = Read-Host "->"
    if($inst_selection -eq 1){
        $valid_selection = $True;
        $install_programs = $True;
        $force_install_programs = $True;
    }elseif($inst_selection -eq 2){
        $valid_selection = $True;
        $install_programs = $True;
        $force_install_programs = $False;
    }elseif($inst_selection -eq 3){
        $valid_selection = $True;
        $install_programs = $False;
        $force_install_programs = $False;
    }
    else{
        Write-Host "Invalid selection"  -foreground "red"
        $valid_selection = $False;
    }
}

# create directory for Classes
$tu_class_dir = "${env:programdata}"+"\tudadesign"
mkdir $tu_class_dir

#---Install Programms---

$isWin64=[System.Environment]::Is64BitOperatingSystem

$install_miktex = $True;
$install_texworks = $True;
if( $force_install_programs -eq $False){
    if( Is-Installed -program "miktex" -eq $True){
        $install_miktex = $False;
        Write-Host "MikteX already installed, not installing"
    }
    if( Is-Installed -program "texworks" -eq $True){
        $install_texworks = $False;
        Write-Host "TeXWorks already installed, not installing"
    }
}

# Installing Miktex
# Downloading Miktex setup
if( $install_miktex -eq $True){
    Write-Host "Installing MikteX"
    if($isWin64 -eq $True){
        Write-Host " 64bit" -NoNewline
        download -url $link_miktex_x64 -dl_name $name_miktex
    }else{
        Write-Host " 32bit" -NoNewline
        download -url $link_miktex_x86 -dl_name $name_miktex
    }
    # unpacking Miktex setup
    Expand-Archive "$name_miktex.zip"

    # downloading and installing Miktex
    miktex-setup/miktexsetup --verbose --local-package-repository=C:\temp\miktex --package-set=complete download;
    miktex-setup/miktexsetup --verbose --local-package-repository=C:\temp\miktex --shared --user-config="<APPDATA>\MiKTeX\2.9" --user-data="<LOCALAPPDATA>\MiKTeX\2.9" --user-install="<APPDATA>\MiKTeX\2.9" --user-roots=$tu_class_dir --print-info-only install;
}else{

write-host "starting mo_admin"  -foreground "yellow"
[console]::beep(500,300)
start-sleep -s 2
mo_admin
write-host "----------------------------------------------"  -foreground "yellow"
write-host "|                  IMPORTANT                  |"  -foreground "yellow" -BackgroundColor "red"
write-host "|                                             |"  -foreground "yellow" 
write-host "| 1. go to the ROOTS tab                      |"  -foreground "yellow"
write-host "| 2. click ADD                                |"  -foreground "yellow"
write-host "| 3. select "("$tu_class_dir\texmf".ToUpper())" |"  -foreground "yellow"
write-host "| 4. click OK                                 |"  -foreground "yellow"
write-host "|                                             |"  -foreground "yellow"
write-host "----------------------------------------------"  -foreground "yellow"
Read-Host "Press any key to continue..."
}

#---Install Classes and Fonts---

# Check if .zip files containing classes and fonts are present
$files_found = $True;
if ((Test-Path "$name_design.zip") -eq $True){
Write-Host "$name_design.zip found";
}else{
Write-Host "$name_design.zip not found";
$files_found = $False;
}

if ((Test-Path "$name_thesis.zip") -eq $True){
    Write-Host "$name_thesis.zip found";
}else{
Write-Host "$name_thesis.zip not found";
$files_found = $False;
}

if ((Test-Path "$name_template.zip") -eq $True){
    Write-Host "$name_template.zip found";
}else{
Write-Host "$name_template.zip not found";
$files_found = $False;
}

if ((Test-Path "$name_fonts.zip") -eq $True){
    Write-Host "$name_fonts.zip found";
}else{
Write-Host "$name_fonts.zip not found";
$files_found = $False;
}

# determine how to proceed if classes and fonts are present
$dl_selection = 2;
if ($files_found -eq $True){
Write-Host "
[1] Use existing files
[2] Redownload files [recommended]"  -foreground "yellow"
$valid_selection = $False;
while($valid_selection -ne $True)
{
$dl_selection = Read-Host "->"
if($dl_selection -ge 1 -and $dl_selection -le 2){
    $valid_selection = $True;
}
else{
    Write-Host "Invalid selection"  -foreground "red"
    $valid_selection = $False;
}
}
}else{
Write-Host "Files incomplete, redownloading"
}

# download classes and fonts if necessary
if ( $dl_selection -eq 2){
download -url $link_design -dl_name $name_design
download -url $link_thesis -dl_name $name_thesis
download -url $link_template -dl_name $name_template
download -url $link_fonts -dl_name $name_fonts
}

Expand-Archive "$name_design.zip" -DestinationPath $tu_class_dir
Write-Host "Unzipping $name_design.zip to $tu_class_dir"
Expand-Archive "$name_thesis.zip" -DestinationPath $tu_class_dir
Write-Host "Unzipping $name_thesis.zip to $tu_class_dir"
Expand-Archive "$name_fonts.zip" -DestinationPath $tu_class_dir
Write-Host "Unzipping $name_fonts.zip to $tu_class_dir"
Expand-Archive "$name_template.zip"
Write-Host "Unzipping $name_template.zip to $pwd"

#-------------STUFF I HAVENT DONE YET JUST COPIED

cd $tu_class_dir
write-host "deleting texmf\fonts\map\dvipdfm"
rmdir /Q /S "texmf\fonts\map\dvipdfm"

Get-ChildItem $env:APPDATA"\MikTeX" | ForEach-Object { 
$path = $_.FullName
write-host "editing "$path"\miktex\config\updmap.cfg"
$nl = [Environment]::NewLine
"Map 5ch.map"+$nl+"Map 5fp.map"+$nl+"Map 5sf.map" | Out-File $path"\miktex\config\updmap.cfg" -encoding ASCII
}

write-host "making maps"
initexmf --mkmaps

#---Testing---

Move-Item -Path Vorlage_Dissertation.pdf -Destination Vorlage_Dissertation_reference.pdf

if ((Get-Command "pdflatex.exe" -ErrorAction SilentlyContinue) -eq $null) 
{ 
   Write-Host "ERROR: Unable to find pdflatex.exe in your PATH" -foreground "red"
}else{
    pdflatex Vorlage_Dissertation.tex
    write-host "----------------------------------------------------------"  -foreground "yellow"
    write-host "| Check $pwd for a Vorlage_Dissertation.pdf file |"  -foreground "yellow"
    write-host "| Compare to Vorlage_Dissertation_reference.pdf          |"  -foreground "yellow"
    write-host "----------------------------------------------------------"  -foreground "yellow"
    Read-Host "Press any key to continue..."
}

write-host "installation complete" -foreground "green"
Read-Host -Prompt "Press Enter to Exit"