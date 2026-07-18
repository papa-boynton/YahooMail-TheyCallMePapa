param(
	[string]$Email,
	[string]$AppPassword
)

$ParameterList = (Get-Command -Name $MyInvocation.InvocationName).Parameters;
    foreach ($key in $ParameterList.keys)
    {
        $var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
        if($var)
        {
            write-host "$($var.name) is $($var.value)"
        }
    }

#py -0p --list-paths

## redirect stderr into stdout
#$p = &{python -V} 2>&1
## check if an ErrorRecord was returned
#$version = if($p -is [System.Management.Automation.ErrorRecord])
##$version = ($p | % gettype) -eq  [System.Management.Automation.ErrorRecord]
#{
#    # grab the version string from the error message
#    $p.Exception.Message
#}
#else 
#{
#    # otherwise return as is
#    $p
#}

#$appName = "Python" 
#if (Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $appName }) { 
#    Write-Host "$appName is installed." 
#} else { 
#    Write-Host "$appName is not installed." 
#}

#python -V

#python.exe YahooMail.py $Email $AppPassword

.\YahooMail $Email $AppPassword