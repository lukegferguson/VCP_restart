#Variables
$LogPath = "C:\OSIS\VCPRegLog.txt"
$LogDir = Split-Path $LogPath
$VCPScript = "$LogDir\Scripts\VCP_Restart.ps1"
$VCPexe = "C:\NextGen\VCP.exe"
$NGDirectory = Split-Path $VCPexe
$Script = $MyInvocation.MyCommand.ScriptContents
$ScriptVer = "0.75"

#Initialize Logging function, stop script and notify user on error
function Write-Log {
    param ($InputObject)

    try {
        if (Test-Path $LogDir){
            Out-File -FilePath $LogPath -Append -InputObject "$(Get-date) $InputObject" -ErrorAction Stop
        } else {
            New-Item -Path $LogDir -ItemType Directory -ErrorAction Stop
            Out-File -FilePath $LogPath -Append -InputObject "$(Get-date) $InputObject" -ErrorAction Stop
         }
    } catch {
        Write-Host "UNABLE TO WRITE LOG FILE" -ForegroundColor red -BackgroundColor White
        Write-host $InputObject
        Write-host "System error: $_" -ForegroundColor red
        Write-host "Ctrl+C to exit" -ForegroundColor Green
        start-sleep -Seconds 100000
    }
}

Write-log "############## Initialize Script Logging ##############"
Write-Log "############## OSIS NextGen VCP Registration Script version $ScriptVer ##############"
Write-Host "OSIS VCP REGISTRATION SCRIPT version $ScriptVer" -foregroundcolor white -backgroundcolor green
Write-Host "Logs: $LogPath"

#Script will attempt to copy itself locally regardless of where it is launched
#Script must be saved for this to work properly, not if pasted into ISE and run directly
function Copy-ScriptLocally {
    $ScriptDir = Split-Path $VCPScript
    Write-Log "Checking for $ScriptDir"
    if (Test-Path $ScriptDir){
        try {
            Out-File -FilePath $VCPScript -Force -InputObject $Script -ErrorAction Stop | Out-Null
            Write-Log "Copied script to $VCPScript"
        } catch {
            Write-Log "Unable to copy script to $VCPScript, copy manually to $ScriptDir before rebooting"
            Write-log "System Error $_"
            write-host "Unable to copy script to $VCPScript, copy manually to $ScriptDir before rebooting" -ForegroundColor Red
            start-sleep -Seconds 30
        }
    } else {
        try {
            New-Item -Path $ScriptDir -ItemType Directory -ErrorAction Stop
            Write-Log "Created $Scriptdir"
            try {
                Out-File -FilePath $VCPScript -Force -InputObject $Script -ErrorAction Stop
                Write-Log "Copied script to $VCPScript"
            } catch {
                Write-Log "Unable to copy script to $VCPScript, copy manually to $ScriptDir before rebooting"
                Write-log "System Error $_"
                write-host "Unable to copy script to $VCPScript, copy manually to $ScriptDir before rebooting" -ForegroundColor Red
            }
        } catch {
            Write-log "Unable to create $VCPScript, copy manually to $ScriptDir before rebooting"
            Write-Log "System Error: $_"
            write-host "Unable to create $VCPScript, copy manually to $ScriptDir before rebooting" -ForegroundColor Red
            start-sleep -Seconds 30
        }
    }
}
Copy-ScriptLocally

#Check Registry for progress tracker, create if it doesn't exist
$RunOnce = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Write-Log "Checking for progress registry entries at $RunOnce"
if ($null -eq (Get-ItemProperty $RunOnce -Name "VCPRegProg" -ErrorAction SilentlyContinue)){

    #Set Progress tracker at 0
    try{
        New-ItemProperty -Path $RunOnce -Name "VCPRegProg" -PropertyType Dword -value 0 -ErrorAction stop
    }
    catch {
        Write-Log "Unable to create registry entry"
        Write-Host "Unable to create registry entry, exiting in 30 seconds" -ForegroundColor Red
        Write-Host "System Error: $_"
        Start-Sleep -Seconds 30
        exit
    }
    
    #Write Logs
    Write-Log  "Progress registry entry created" 
    Write-Log  "VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"

} else {
    
    Write-Log "VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"

}

#Initialize script variable from registry entry
$Progress = (Get-ItemProperty $RunOnce -Name "VCPRegProg").VCPRegProg

#Function to start VCP process
#-Registration $true causes /r switch to be used
#-ZipDelete $True will delete all ZIP files in $NGDirectory
function Start-VCP {
    param(
        $Register = $false,
        $ZipDelete = $False
    )

    Write-Log "Starting $VCPexe Register=$Register, ZipDelete=$ZipDelete"

    if (Test-Path -Path $VCPexe){
        Write-log "EXE exists: $VCPexe"
    }
    else {
        Write-log "$VCPexe not found"
        Write-Host "Executable not found at $VCPexe" -ForegroundColor Red
        Write-Host "Exiting Script in 30 seconds"
        Start-Sleep -Seconds 30
        exit
    }

    if ($ZipDelete -eq $true){  
        # VCP Registration process attempts to verify files exist, however it ony checks for the ZIPs, not the unpacked files
        #Deleting the ZIP files prior to running VCP registration forces it to re-download the ZIPs and unpack their contents, hopefully correcting anything missing/corrupt
        $ZIPs = Get-ChildItem $NGDirectory | Where-Object -Property Name -like *.zip
        Write-Log -InputObject ".ZIP files contained in ($NGDirectory): $ZIPs"
       try { 
        $ZIPs | Remove-Item -ErrorAction Stop
        Write-Log -InputObject "Operation to delete ZIPs complete"
        } catch {
            Write-Log "Operation to delete ZIPs incomplete"
            Write-log "System Error $_"
        }
    }

    #log .NET Installed before changes, as it may be changed during VCP registration
    try {
        Write-Log ".NET Runtimes currently installed: $(dotnet --list-runtimes)" -ErrorAction Stop
    }
    catch {
        Write-Log "dotnet --list-runtimes not returning, likely none installed"
        Write-Log "System Error: $_"
    }

    #Stop any NextGen processes before proceeding
    $ProcToStop = get-process NG*
    Write-Log "NextGen processes running: $($ProcToStop.Name)"
    foreach ($Proc in $ProcToStop){
        try {
            Stop-Process -Name $proc.Name -ErrorAction stop
            Write-Log "Stopped process $($Proc.name)"
        }
        catch {
            Write-Log "Unable to stop process $($Proc.name) $($Error)"
        }
    }

    #Switch user mode to allow for software install, changes how Windows handles INI files
    $install = change user /install
    Write-Log -InputObject "$install"

    #VCP Registration process start based on /r switch
    if ($registration) {
        Write-Log -InputObject "Starting $VCPexe /r"
        try { Start-Process -FilePath $VCPexe -ArgumentList "/r" -ErrorAction stop } 
        catch {
        Write-log "System Error: $_"
        Write-Log "Exiting script"
        exit
    }
    } else {
        Write-Log -InputObject "Starting $VCPexe no switches"
        try  { Start-Process -FilePath $VCPexe -ErrorAction Stop }
        catch { 
        Write-log "System Error: $_"
        Write-Log "Exiting script"
        exit
    }
}
    do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))

    Write-Log -InputObject "VCP Process Exited"    
}

#Add Script Location to RunOnce to execute again after reboot
function Start-NextStep {

    param(
        $Reboot = $false
    )

    Write-Log "Starting Next Script Step, Reboot: $reboot"

    try {
        New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "Powershell.exe $VCPScript" -ErrorAction Stop
        Write-Log -InputObject "Script location added to registry RunOnce: $VCPScript" 
    } catch {
        Write-Log "Unable to write script location to $runonce"
        Write-log "System error: $_"
        exit
    }

    #Increase progress count, set registry key
    $Script:Progress += 1
    try { 
        Set-ItemProperty -Path $RunOnce -Name "VCPRegProg" -Value $($progress) -ErrorAction stop
        Write-Log "Script will resume when an administrator logs in"
        Write-Log "Updated progress in registry to $progress"

        #Notify and reboot
        if ($Reboot) {
            Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
            Start-Sleep -Seconds 15
            Restart-Computer -Force
        }
    }
    catch {
        Write-Log "Unable to write progress in registry"
        Write-log "System error: $_"
        exit
    }
}

#Change script behavior based on progress tracking registry entry
switch ($progress){
    #Execute this section on first and second script run
    {($_ -le 1)} {
        Start-VCP -Register $true -ZipDelete $true
        Start-NextStep -Reboot $true
    }

    #Run on third iteration of script
    (2) {
        1..4 | foreach { Start-VCP }
        Start-NextStep -Reboot $true
    }

    #Fourth iteration of script
    (3) {
        #Switch user mode back to normal
        $execute = change user /execute
        Write-Log -InputObject "$execute"
        
        #Delete progress tracking registry entry
        #RunOnce scripts are removed upon being run, so no need to remove
        try {
            Remove-ItemProperty -Path $RunOnce -Name "VCPRegProg" -ErrorAction Stop
            Write-Log "Removed registry for progress tracking"
        } catch {
            Write-Log "Unable to remove tracking registry entry"
        }
        
        #Finalize Logs
        Write-Log -InputObject "##### VCP Registration Complete #####"
    }
        
    default {
        Write-Log -InputObject "real big oops there 😣"
    }
}