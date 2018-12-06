## Check LabTech agent and re/install if needed.
## Written by: Alex Heylin @ Gregory Micallef Associates - www.gmal.co.uk  
## Inspired by batch file written by darrenwhite99 @ LTG Slack - thanks for sharing Darren!
## Requires (and dynamically imports) the Labtech-Powershell-Module
## https://github.com/LabtechConsulting/LabTech-Powershell-Module
## Many thanks to http://labtechconsulting.com for their excellent work
## License on my code: MIT license - please leave header intact and consider submitting improvements to this project
## via the public repo: https://github.com/AlexHeylin/Check-LTAgent
## Supplied with no warranty etc. If your cat dies - don't blame me. 
## Change log: 2018-11-07 Initial version.
## 2018-11-14 Tidy up comments and code. Improve flow control logic.
## 2018-11-15 Rework so we can invoke this directly from scheduled task via HTTPS.
## 2018-11-15 First public release on Github - Enjoy!  (I know it's messy, but seems to work.)
## 2018-12-06 Rework the services section to be more forgiving and retry changes to avoid unnecessary reinstalls.


## If you want to set default / override values, do that here
# $LTSrv = "labtech.mymspname.here"
# $LogFile = $null
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

  ## Check TerminalServerMode using WMI.  1 = Application Server mode (Terminal Server), 0 = Remote Administration mode (normal RDP)
  if ( (Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -Class "Win32_TerminalServiceSetting" ).TerminalServerMode -eq 1) {
    outlog "Terminal server detected. Switching it to INSTALL mode"
    $result = iex "$env:windir\system32\CHANGE USER /INSTALL"
    outlog "Result was $result"
  }

  ## Install .Net 3.5 if missing
  $Result = Dism /online /get-featureinfo /featurename:NetFx3
  If($Result -contains "State : Enabled")	{
    outlog ".Net Framework 3.5 is installed and enabled."
  } Else {
      outlog ".Net Framework 3.5 not detected, calling DISM to install"
      Dism /online /Enable-feature /featurename:NetFx3 | Out-Null 
      $Result = Dism /online /Get-featureinfo /featurename:NetFx3
      If($Result -contains "State : Enabled"){
        outlog "Installed .Net Framework 3.5 successfully."
      } Else {
        outlog "Failed to install install .Net Framework 3.5. Result[$result]"
          outlog "Trying to install LabTech agent as we've got nothing to lose."
      }
  }

  ## Dynamically load the latest LT PoSh module
	outlog "Dynamically loading LT-PoSh module from Github"
	(new-object Net.WebClient).DownloadString('http://bit.ly/LTPoSh') | iex ; 
	outlog "Calling Reinstall-LTService -Server $LTSrv -LocationID $LTLoc"
	## Call the module to do the (re)install
	#Reinstall-LTService -Server $LTSrv -LocationID $LTLoc
	Reinstall-LTService -Server $LTSrv -LocationID $LTLoc >> $LogFile

	## Check TerminalServerMode using WMI.  1 = Application Server mode (Terminal Server), 0 = Remote Administration mode (normal RDP)
	if ( (Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -Class "Win32_TerminalServiceSetting" ).TerminalServerMode -eq 1) {
		outlog "Terminal server detected. Switching it to EXECUTE mode"
		$result = iex "$env:windir\system32\CHANGE USER /EXECUTE"
		outlog "Result was $result"
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
	    outlog "Registry checks OK. It seems LT is installed - checking that services look OK."
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
