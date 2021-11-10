###########################################
#                                         #
#   Agent installation example script     #
#                                         #
###########################################

$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'

$ServerList = "Servers.Csv"

if(!Test-Path $ServerList) { 
    Write-Host "Unable to open $ServerList" 
    Write-Host -NoNewLine "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "Obtaining list of computers $ServerList....`n"

$objServers = Import-Csv -Path $ServerList

# Check that script is running as an administrator
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        Write-Host "Please run this script as an administrator`n"
        Write-Host -NoNewLine "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

Write-Host "Prompting for credentials for pre-authentication...`n"
             
do {
    $Auth = $False
    $cred = Get-Credential -Credential $null
    $ConResult = Invoke-Command -ComputerName localhost -Credential $cred -ScriptBlock { Write-Output $Env:ComputerName } -ErrorVariable Err
    if ($Err.Count -gt 0) {
        If ($Err.Exception.TransportMessage -like "*incorrect*" -or $Err.Exception.TransportMessage -like "*denied*") {
            Write-Host "Incorrect user name or password"
            Read-Host "Hit enter to try again or <Ctrl><C> to exit"
        }
        else {
            Write-Host -ForegroundColor red "Unable to validate credentials : $($Err.Exception.TransportMessage)"
            Break
        }
    }
    else {
        Write-Verbose "Pre-authentication was successful!"
        $Auth = $True
    }
}
Until ($Err.Count -eq 0)

if($Auth -eq $True -and $cred) { 
    
    foreach ($Server in $objServers) {
        $ComputerName = $Server.Servers

        Write-Host "Installing Unitrends Agent on $ComputerName....`n"
        Invoke-Command -ComputerName $ComputerName -Credential $cred -ScriptBlock { Start-Process -Wait -FilePath msiexec.exe -ArgumentList /i,https://bpagent.s3.amazonaws.com/latest/windows/Unitrends_Agentx64.msi,/norestart,/qn }

        if ($(Invoke-Command -ComputerName $ComputerName -Credential $cred -ScriptBlock { Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -ilike 'Unitrends Agent*' } | Select DisplayName })) {
            Write-Host "Verified that the Unitrends Agent is installed"
           
            Write-Host "Installing CBT driver on $ComputerName....`n"
            Invoke-Command -ComputerName $ComputerName -Credential $cred -ScriptBlock { Start-Process -Wait -FilePath msiexec.exe -ArgumentList /i,C:\PCBP\Installers\uvcbt.msi,/norestart,/qn } 
        
            if ($(Invoke-Command -ComputerName $ComputerName -Credential $cred -ScriptBlock { Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -ilike 'Unitrends Volume CBT Driver*' } | Select DisplayName })) {
                Write-Host "Verified that the CBT Driver is installed"
            } else 
            {
                Write-Host "Failed to install CBT driver on $ComputerName"
            }
        } else 
        {
            Write-Host "Failed to install Unitrends Agent on $ComputerName"
        }    
      Write-Host  
    }
}

Write-Host -NoNewLine "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")