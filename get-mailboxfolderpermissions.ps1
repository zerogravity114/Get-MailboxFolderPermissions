# Get-Mailboxpermissionreport


$365creds = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $365creds -Authentication Basic -AllowRedirection

Import-PSSession $Session | Out-Null

$Mailboxes = Get-Mailbox

$MailboxCount = $Mailboxes.Count
$Summary = @();
$i=0;
For ($i -eq 0; $i -lt $MailboxCount; $i++) {
    $Mailbox = $Mailboxes[$i]
    $Email = $Mailbox.PrimarySmtpAddress
    Write-Warning "Processing $Email"
    [string]$CalendarPath = $Email + ":\Calendar"
    $AccessRights = Get-MailboxFolderPermission -Identity $CalendarPath
    ForEach ($Right in $AccessRights) {
    [string]$User = $Right.User
    [string]$Permission = $Right.AccessRights
    $object = New-Object -TypeName PSObject
    $object | Add-Member -MemberType NoteProperty -Name Mailbox -Value $Email
    $object | Add-Member -MemberType NoteProperty -Name User -Value $User
    $object | Add-Member -MemberType NoteProperty -Name AccessRight -Value $Permission
    $Summary += $object

    }


}
