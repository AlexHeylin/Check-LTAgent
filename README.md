# Check-LTAgent

Aims to check the LabTech agent on a Windows 7+ machine seems operable and re/install if required.  
Designed so it can be directly called from a scheuled task (possible created by Group Policy).

I welcome improvments to this code, please submit as pull requests.

## Sample usage

$LTLOCATIONID = 1 ; (new-object Net.WebClient).DownloadString('/Check-LTAgent/Check-LTAgent.ps1') | iex ;
