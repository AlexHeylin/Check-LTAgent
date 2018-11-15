# Check-LTAgent
Aims to check the LabTech agent on a Windows 7+ machine seems operable and re/install if required.  
Designed so it can be directly called from a scheuled task (possible created by Group Policy).

Written in Powershell for simplicity.
Aims to provide compatability from Windows 7 - 10 (Server 2008 - 2019), but testing has been limited to please report any issues
I welcome improvments to this code, please submit as pull requests or if you can't do that, post details of problem and fix as an Issue.


## Sample usage
$LTLOCATIONID = 1 ; (new-object Net.WebClient).DownloadString('/Check-LTAgent/Check-LTAgent.ps1') | iex ;
