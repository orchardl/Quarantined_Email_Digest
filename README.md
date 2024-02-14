Hello there,

I didn't like Microsoft's built-in tool to send users a notification when they have an email to them that is stuck in quarantine, so I built my own!

If you need help setting up an App and Certification to connect Exchange Online, this worked like a charm for me:
https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps

If you're going to run this as a different user, I had to set up the certificate using,

<code>PS> Enter-PSSession localhost -Credential (Get-Credential)</code>

and from there I set up the certificate on the machine.

This is another good resource to get the permissions you need for this task: https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps I gave it the "Security Reader" role, and I haven't had any issues.

Also, make sure you've run,

<code>PS> Install-Module ExchangeOnlineManagement</code>

beforehand on the machine.
