$sessions = Get-ChildItem 'HKCU:\Software\SimonTatham\PuTTY\Sessions'
#$dir = ("{0}\Microsoft\Windows\Start Menu\Programs\ssh" -f $env:APPDATA)
$dir = ("{0}\ssh_connections" -f $HOME)

rm $dir -Confirm -ErrorAction SilentlyContinue
mkdir $dir -ErrorAction SilentlyContinue 

foreach ($i in $sessions) {
    
    $name       = ($i.Name.Substring($i.Name.LastIndexOf("\") + 1)).Replace("%20"," ")
        
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut(("{0}\ssh_{1}.lnk" -f $dir,$name))
    $Shortcut.TargetPath = ("{0}\putty\putty.exe" -f $HOME)
    $Shortcut.Arguments = ("-load {0}" -f $name)
    $Shortcut.Save()
}

