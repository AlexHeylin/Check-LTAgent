## Check LabTech agent and re/install if needed.
## Written by: Alex Heylin @ Gregory Micallef Associates - www.gmal.co.uk  
## Inspired by batch file written by darrenwhite99 @ LTG Slack - thanks for sharing Darren!
## Requires (and dynamically imports) the Labtech-Powershell-Module
## https://github.com/LabtechConsulting/LabTech-Powershell-Module
## Many thanks to http://labtechconsulting.com for their excellent work
## License on my code: MIT license - please leave header intact and consider submitting improvements to this project
## via the public repo: https://github.com/AlexHeylin/Check-LTAgent
## Supplied with no warranty etc. If your cat dies - don't blame me. 
## Be aware this code may block, and take up to 5-10 minutes to run. Do not use as a user login script!
## Change log: 2018-11-07 Initial version.
## 2018-11-14 Tidy up comments and code. Improve flow control logic.
## 2018-11-15 Rework so we can invoke this directly from scheduled task via HTTPS.
## 2018-11-15 First public release on Github - Enjoy!  (I know it's messy, but seems to work.)
## 2018-12-06 Rework the services section to be more forgiving and retry changes to avoid unnecessary reinstalls.


## If you want to set default / override values, do that here
## $LTSrv = "labtech.mymspname.here"
## $LogFile = "c:\windows\temp\Check-LTAgent.txt"
## Default to location ID 1 (default in LT)
If ($LTLoc -eq $null) {$LTLoc = 1}

### You should not need to modify below here ###

## Functions (at the top so it can be debugged easily)

function outlog {
    Param($LogLine)
    write-output "$LogLine"
    if ($LogFile -ne "" -and $LogFile -ne $null) {
        $FileLine = (get-date -Format 'yyyy-MM-dd HH:mm:ss') + " $LogLine"
        Add-Content $LogFile $FileLine
    }
}

Function Reinstall {
	outlog "Starting Reinstall function"

	outlog "Trying to repair LTAgent first..."
	## Dynamically load the latest LT PoSh module
	outlog "Dynamically loading LT-PoSh module from Github"
	try {
		(new-object Net.WebClient).DownloadString('http://bit.ly/LTPoSh') | iex ; 
	} catch {
		$ErrorMessage = $_.Exception.Message
		outlog "EXCEPTION: $ErrorMessage"
		throw
	}

	if (((('LTService') | Get-Service -EA 0 | Measure-Object | Select-Object -Expand Count) -eq 0) -or ((('LTSvcMon') | Get-Service -EA 0 | Measure-Object | Select-Object -Expand Count) -eq 0)) {
		outlog "Services found - Calling Restart-LTService"
		try {
			$startResult = Restart-LTService
			## NOTE: Do NOT rely on the text of the result. It can lie and say services started when they LTService did not start. 
			## TODO: Try and fix & submit change to LT-PoSh for this.
		} catch {
			$ErrorMessage = $_.Exception.Message
			outlog "EXCEPTION: $ErrorMessage"	
			outlog "Result of Restart-LTService was $startResult"
			throw
		}
	} else {
		outlog "At least one service not found"
	}
	
	if (((Get-Service LTService -EA 0 -WA 0 ).Status -ne "Running") -or ((Get-Service LTSvcMon -EA 0 -WA 0 ).Status -ne "Running")) {
		outlog "Services didn't start properly, proceeding with reinstall"
		## Check TerminalServerMode using WMI.  1 = Application Server mode (Terminal Server), 0 = Remote Administration mode (normal RDP)
		if ( (Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -Class "Win32_TerminalServiceSetting" ).TerminalServerMode -eq 1) {
			outlog "Terminal server detected. Switching it to INSTALL mode"
			$result = iex "$env:windir\system32\CHANGE USER /INSTALL"
			outlog "Result was $result"
		}


		## OK, we're going to have to reinstall.
		outlog "We're going to have to get more brutal and try a (re)install. Taking backup of settings first"
		$backupResult = New-LTServiceBackup
		outlog "Result of New-LTServiceBackup was $backupResult"

		## Use any existing LocationID unless $forceLocID set
		## This allows agents to REinstall to the location they're already in LT, even if cmd line params use a different default site.
		## Stops agents moving back to "main" location if they've been manually moved to another location in LT.
		If (($forceLocID -ne $true) -and ((get-LTServiceInfo).LocationID -ge 1)){
			outlog "Found existing LocationID and forceLocID not set so using it"
			$LTLoc = (get-LTServiceInfo).LocationID
		}

		outlog "Calling Reinstall-LTService -Server $LTSrv -LocationID $LTLoc"
		## Call the module to do the (re)install
		#Reinstall-LTService -Server $LTSrv -LocationID $LTLoc
		$InstallResult = Reinstall-LTService -Server $LTSrv -LocationID $LTLoc >> $LogFile
		outlog "Result of Reinstall-LTService was $InstallResult"

		## Check TerminalServerMode using WMI.  1 = Application Server mode (Terminal Server), 0 = Remote Administration mode (normal RDP)
		if ( (Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -Class "Win32_TerminalServiceSetting" ).TerminalServerMode -eq 1) {
			outlog "Terminal server detected. Switching it to EXECUTE mode"
			$result = iex "$env:windir\system32\CHANGE USER /EXECUTE"
			outlog "Result was $result"
		}
	} else {
		outlog "Services started OK so we avoided reinstall :-)"
	}
	outlog "Finished reinstall function"
}


## Main Flow
If ($LTSrv -eq "labtech.mymspname.here" -or $LTSrv -eq "" -or $LTSrv -eq $null) { 
	outlog "You need to specify the LT server to use by setting `$LTSrv variable before calling this script" ;	
} else {

    outlog "Checking health of LabTech agent" ;

    if ( ((Get-ItemProperty -Path HKLM:\SOFTWARE\LabTech\Service -Name "ID" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).ID -eq $null) `
		    -or ((Get-ItemProperty -Path HKLM:\SOFTWARE\LabTech\Service\Settings  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).ServerAddress -notlike "*" + $LTSrv)
	    )  
    { 
        outlog "The LT agent registry keys do not look right - going to reinstall"
	    Reinstall
    } else {
	    outlog "Registry checks OK. Checking services exist"
		if ((('LTService') | Get-Service -EA 0 | Measure-Object | Select-Object -Expand Count) -eq 0) {
			outlog "LTService is missing. Calling reinstall."
			Reinstall
		} else {
			outlog "LTservice is installed"
		}

		if ((('LTSvcMon') | Get-Service -EA 0 | Measure-Object | Select-Object -Expand Count) -eq 0) {
			outlog "LTSvcMon is missing. Calling reinstall."
			Reinstall
		} else {
			outlog "LTSvcMon is installed"
		}		
		outlog "It seems LT is installed - checking that services look OK."
		
	    outlog "Checking LTService is set to Auto start"
		If ((Get-Service LTService -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).StartType -ne "Automatic") { 
			outlog "LTService is not set to Auto start. Attempting to set it to Auto"
			Try {
				Set-Service LTService -startupType Automatic -ErrorAction Stop -WarningAction Stop 
				outlog "OK"
			} catch {
				$ErrorMessage = $_.Exception.Message
				outlog "EXCEPTION: $ErrorMessage"
				outlog "Waiting 120 seconds and trying again"
				Start-Sleep 120 
				Try {
					Set-Service LTService -startupType Automatic -ErrorAction Stop -WarningAction Stop 
					outlog "OK"
				} catch {
					$ErrorMessage = $_.Exception.Message
					outlog "EXCEPTION: $ErrorMessage"
					outlog "Service LTService is broken. Calling Reinstall"
					Reinstall
				}
			}
		} else {
			outlog "LTService already set to Auto start"
		}

		outlog "Checking LTService is Running"
		If ((Get-Service LTService -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Status -ne "Running") { 
			outlog "LTService is not Running. Attempting to start it."
			Try {
				Start-Service LTService -ErrorAction Stop -WarningAction Stop  
			} catch {
				$ErrorMessage = $_.Exception.Message
				outlog "EXCEPTION: $ErrorMessage"
				outlog "Waiting 120 seconds and trying again"
				Start-Sleep 120 
				Try {
					Start-Service LTService -ErrorAction Stop -WarningAction Stop  
				} catch {
					$ErrorMessage = $_.Exception.Message
					outlog "EXCEPTION: $ErrorMessage"
					outlog "Service LTService is would not start. Calling Reinstall"
					Reinstall
				}
			}
		} else {
			outlog "LTService already Running"
		}

		outlog "Checking LTSvcMon is set to Auto start"
		If ((Get-Service LTSvcMon -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).StartType -ne "Automatic") { 
			outlog "LTSvcMon is not set to Auto start. Attempting to set it to Auto"
			Try {
				Set-Service LTSvcMon -startupType Automatic -ErrorAction Stop -WarningAction Stop 
				outlog "OK"
			} catch {
				$ErrorMessage = $_.Exception.Message
				outlog "EXCEPTION: $ErrorMessage"
				outlog "Waiting 120 seconds and trying again"
				Start-Sleep 120 
				Try {
					Set-Service LTSvcMon -startupType Automatic -ErrorAction Stop -WarningAction Stop 
					outlog "OK"
				} catch {
					$ErrorMessage = $_.Exception.Message
					outlog "EXCEPTION: $ErrorMessage"
					outlog "Service LTSvcMon is broken. Calling Reinstall"
					Reinstall
				}
			}
		} else {
			outlog "LTSvcMon already set to Auto start"
		}

		outlog "Checking LTSvcMon is Running"
		If ((Get-Service LTSvcMon -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Status -ne "Running") { 
			outlog "LTSvcMon is not Running. Attempting to start it."
			Try {
				Start-Service LTSvcMon -ErrorAction Stop -WarningAction Stop
				outlog "OK"  
			} catch {
				$ErrorMessage = $_.Exception.Message
				outlog "EXCEPTION: $ErrorMessage"
				outlog "Waiting 120 seconds and trying again"
				Start-Sleep 120 
				Try {
					Start-Service LTSvcMon -ErrorAction Stop -WarningAction Stop  
					outlog "OK"
				} catch {
					$ErrorMessage = $_.Exception.Message
					outlog "EXCEPTION: $ErrorMessage"
					outlog "Service LTSvcMon is would not start. Calling Reinstall"
					Reinstall
				}
			}
		} else {
			outlog "LTSvcMon already Running"
		}
		outlog "Labtech Agent checks completed OK. Enjoy the rest of your uptime!"
    }
}
