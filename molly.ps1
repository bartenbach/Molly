# Molly's Laptop Setup Checklist
# Version 0.0.2
# Blake Bartenbach
$version = '0.0.3'

function init {
  write-host "[Molly's Laptop Setup Checklist $version]" -foregroundcolor "yellow"
  write-host ""
}

function apply-Power-Settings {
  write-host "[Setting High Performance Power Scheme]" -foregroundcolor "green"
  powercfg -s SCHEME_MIN >> $null
  write-host "[Disabling AC USB Selective suspend]" -foregroundcolor "green"
  powercfg -setacvalueindex SCHEME_MIN 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 000
  write-host "[Disabling DC USB Selective suspend]" -foregroundcolor "green"
  powercfg -setdcvalueindex SCHEME_MIN 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 000
}

function get-Windows-Updates {
  write-host "[Checking for Windows Updates]" -foregroundcolor "green"
  wuapp.exe
  wuauclt.exe /updatenow
}

function disableUsbHubPowerSaving {
  $hubs = Get-WmiObject Win32_USBHub
  $powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi
  foreach ($p in $powerMgmt) {
	  $IN = $p.InstanceName.ToUpper()
	  foreach ($h in $hubs) {
	    $PNPDI = $h.PNPDeviceID
      if ($IN -like "*$PNPDI*") {
	      write-host "[USB HUB power saving feature disabled]" -foregroundcolor "green"
        $p.enable = $False
        $p.psbase.put() >> $null
      }
	  }
  }
}

function disableNetAdapterPowerSaving { 
  $PhysicalAdapters = Get-WmiObject -Class Win32_NetworkAdapter|Where-Object{$_.PNPDeviceID -notlike "ROOT\*" `
      -and $_.Manufacturer -ne "Microsoft" -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22} 
	
  foreach($PhysicalAdapter in $PhysicalAdapters) {
    $PhysicalAdapterName = $PhysicalAdapter.Name
    $DeviceID = $PhysicalAdapter.DeviceID
    if([Int32]$DeviceID -lt 10) {
      $AdapterDeviceNumber = "000"+$DeviceID
    } else {
      $AdapterDeviceNumber = "00"+$DeviceID
    }
		
    $KeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\$AdapterDeviceNumber"
    if(Test-Path -Path $KeyPath) {
      $PnPCapabilitiesValue = (Get-ItemProperty -Path $KeyPath).PnPCapabilities
      if($PnPCapabilitiesValue -eq 24) {
        write-host "[""$PhysicalAdapterName"" power saving feature already disabled]" -foregroundcolor "green"
      }
      if($PnPCapabilitiesValue -eq 0) {
        try {	
          #setting the value of properties of PnPCapabilites to 24, it will disable save power option.
          Set-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 | Out-Null
          write-host "[""$PhysicalAdapterName"" power saving feature disabled]" -foregroundcolor "green"
        } catch {
          write-host "[Failed to disable power saving of network adapter]" -foregroundColor "red"
        }
      }
      if($PnPCapabilitiesValue -eq $null) {
        try {
          New-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 -PropertyType DWord | Out-Null
		      write-host "[""$PhysicalAdapterName"" power saving feature disabled]" -foregroundcolor "green"
        } catch {
		      write-host "[Failed to disable power saving of network adapter]" -foregroundcolor "red"
	      }
	    }
    } else {
			Write-Warning "The path ($KeyPath) not found."
		}
	}
}

function reboot-computer {
	write-host "It will take effect after reboot, do you want to reboot right now? (y/n): " -foregroundColor "yellow" -noNewline
	[string]$Reboot = Read-Host
	if ($Reboot -eq "y" -or $Reboot -eq "yes") {
		Restart-Computer -Force
    Exit
	}
}

function rename-computer {
	$friendlyName = Get-WmiObject Win32_operatingsystem | select-object -expand csname
	$computerName = Get-WmiObject -Class Win32_ComputerSystem
	write-host "Current computer name is: " -foregroundColor "yellow" -noNewline
	write-host $friendlyName -foregroundColor "green"
	write-host "Would you like to rename? (y/n): " -foregroundColor "yellow" -noNewline
	[string]$prompt = read-host

	if ($prompt -eq "y" -or $prompt -eq "yes") {
	  write-host "New Name: " -foregroundColor "yellow" -noNewline
	  [string]$newname = read-host
    # TESTING
		$success = $computername.Rename($newname)
    if ($success.ReturnValue -eq 0) {
      Write-Host "Computer renamed successfully" -ForegroundColor "Green"
      write-host ""
		  reboot-computer
    } else {
      Write-Host "Failed to rename computer" -foregroundColor "Yellow"
      Write-Host ""
      Write-Host "Adding computer to Workgroup A" -foregroundColor "Green"
      Add-Computer -WorkgroupName A
      # can we rename the computer here?
      Write-Host "Attempting to rename computer..." -ForegroundColor "Green"
      $success = $computername.Rename($newname)
      Write-Host "Attempting to re-join domain..." -ForegroundColor "Green"
      Add-Computer -DomainName chsomaha.org
      Write-Host "(Not sure if this actually works...)" -ForegroundColor "Yellow"
      reboot-computer
    }

	} else {
    get-Windows-Updates
  }
}

function main {
  init
  apply-Power-Settings
  disableNetAdapterPowerSaving
  disableUsbHubPowerSaving
  write-host ""
  rename-computer
  elephant
  write-host "Done!" -foregroundcolor "green"
  write-Host "Press any key to continue ..."
  $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function elephant {
  write-host " __                                    " -foregroundcolor white
  write-host "'. \                                   " -foregroundcolor white
  write-host " '- \                                  " -foregroundcolor white
  write-host "  / /_         .---.                   " -foregroundcolor white
  write-host " / | \\,.\/--.//    )                  " -foregroundcolor white
  write-host " |  \//        )/  /                   " -foregroundcolor white
  write-host "  \  ' ^ ^    /    )____.----..  6     " -foregroundcolor white
  write-host "   '.____.    .___/            \._)    " -foregroundcolor white
  write-host "      .\/.                      )      " -foregroundcolor white
  write-host "       '\                       /      " -foregroundcolor white
  write-host "       _/ \/    ).        )    (       " -foregroundcolor white
  write-host "      /#  .!    |        /\    /       " -foregroundcolor white
  write-host "      \  C// #  /'-----''/ #  /        " -foregroundcolor white
  write-host "   .   'C/ |    |    |   |    |mrf  ,  " -foregroundcolor white
  write-host "   \), .. .'OOO-'. ..'OOO'OOO-'. ..\(, " -foregroundcolor white
  write-host ""
}

main