# AppointmentRetriever

This console application is an example which demonstrates how one can retrieve all resource mailboxes designated as a meeting room from Active Directory and then retrieve each of their appointments.

For this to work (view the calendar of a mailbox from another user), a user is required that has Full Access rights to each of the room mailboxes you wish to access. By connecting to the exchange server via a powershell session, using an admin account, we can issue give a user this level of access rights to all room mailboxes. 

* A note on FullAccess: Full access “Allows the delegate to open the mailbox, and view, add and remove the contents of the mailbox. Doesn't allow the delegate to send messages from the mailbox.”

#### The follow PowerShell commands will issue FullAccess rights to a user

```powershell
$URL = "<url to your exchange server>"
$Creds = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Authentication Kerberos -Credential $Creds
Import-PSSession $Session
Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'RoomMailbox') -and (Alias -ne 'Admin')} | Add-MailboxPermission -User <the user email you want to assign access to> -AccessRights fullaccess -InheritanceType all -AutoMapping:$false
```

####Other userful scripts

Get all room mailboxes
```powershell
Get-Mailbox -Filter {RecipientTypeDetails -eq "RoomMailbox"}
```

Get permissions on a room for a user
```powershell
Get-MailboxPermission -Identity <room mailbox (email)> -User "<user alias>"
```
