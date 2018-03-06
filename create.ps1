
function findPutty(){
    $search = Get-ChildItem -Path C:\ -Filter putty.exe -Recurse -ErrorAction SilentlyContinue |  Where-Object { $_.Attributes -ne "Directory"} | Select-Object -First 1
    if ( $? -eq $true) {
        return ("{0}\{1}" -f $search.DirectoryName,$search.Name)
    }
    else{
        return $false
    }
}

$sessions = Get-ChildItem 'HKCU:\Software\SimonTatham\PuTTY\Sessions'
#$dir = ("{0}\Microsoft\Windows\Start Menu\Programs\ssh" -f $env:APPDATA)
$dir = ("{0}\ssh_connections" -f $HOME)
$puttypath = findPutty
if ($puttypath -eq $false ){
    echo "No putty on c:\`nEnter to exit..."
    Read-Host
    exit 2
}

rm $dir -Recurse -ErrorAction SilentlyContinue | Out-Null
mkdir $dir -ErrorAction SilentlyContinue  | Out-Null

foreach ($i in $sessions) {
    
    $name = ($i.Name.Substring($i.Name.LastIndexOf("\") + 1)).Replace("%20"," ")
        
    $WshShell = New-Object -comObject WScript.Shell
    # Commented out - with ssh prefix
    #$Shortcut = $WshShell.CreateShortcut(("{0}\ssh_{1}.lnk" -f $dir,$name))
    $Shortcut = $WshShell.CreateShortcut(("{0}\{1}.lnk" -f $dir,$name))
    $Shortcut.TargetPath = $puttypath
    $Shortcut.Arguments = ("-load {0}" -f $name)
    $Shortcut.Save()
}

