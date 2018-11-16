# Check-LTAgent
Aims to check the LabTech agent on a Windows 7+ machine seems operable and re/install if required using specified LT server and location ID.  
Designed so it can be directly called from a scheuled task (possibly created by Group Policy).

Written in Powershell for simplicity.
Aims to provide compatability from Windows 7 - 10 (Server 2008 - 2019), but testing has been limited so please report any issues.

I welcome improvments to this code, please submit as pull requests or if you can't do that, post details of problem and fix as an Issue.

## Variables (yes, they're not parameters - this seemed easier)
This code accepts you setting some variables before execution
```powershell
$LTSrv = "labtech.mymspname.here" # FQDN of your LabTech server (required - unless you modify script header)
$LTLoc = 1 # LabTech Location ID to add agent to (optional - defaults to 1)
$LogFile = "$Env:temp\Check-LTAgent-log.txt" # If you want a logfile, where do you want the file. (optional)
```

## Sample usage
```powershell
$LTSrv = 'labtech.mymspname.here' ; $LTLoc = 1 ;  (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/AlexHeylin/Check-LTAgent/master/Check-LTAgent.ps1') | iex ;
```

or in short form (if the Bitly stays live)
```powershell
$LTSrv='labtech.mymspname.here';$LTLoc=1;(new-object Net.WebClient).DownloadString('https://bit.ly/2qO49e8')|iex;
```
Watch your quote types " vs ' when calling this directly from CMD.

Might need to use this on older versions - need to check compatability
```powershell
$LTSrv='labtech.mymspname.here';$LTLoc=1;(new-object System.IO.StreamReader((([System.Net.WebRequest]::Create('https://bit.ly/2qO49e8')).GetResponse()).GetResponseStream())).ReadToEnd()|iex;
```
