$Threshold = 90
$Today = Get-Date
$Days = New-TimeSpan -Days $Threshold
$Output = @()
$Outpath = 'C:\Temp\InactiveOffice365.csv'
$AllUserMailboxes = Get-Mailbox -ResultSize Unlimited |
    Where-Object { $_.WhenMailboxCreated -lt ($Today - $Days) }
$AllUserMailboxesStats = $AllUserMailboxes |
    Get-MailboxStatistics -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
$InactiveMailboxes = $AllUserMailboxesStats |
    Where-Object { ($_.LastLogonTime -lt ($Today - $Days)) -or ($_.LastLogonTime -eq $Null) }
foreach ($InactiveMailbox in $InactiveMailboxes) {
    $Mailbox = ($AllUserMailboxes |
            Where-Object { $_.DisplayName -eq $InactiveMailbox.DisplayName })

    if ($Null -eq $InactiveMailbox.LastLogonTime) {
        $DaysInactive = ($Today - $Mailbox.WhenMailboxCreated).days
    } else {
        $DaysInactive = ($Today - $InactiveMailbox.LastLogonTime).days
    }

    $Instance = ($InactiveMailbox |
            Select-Object @{
                Label = 'Displayname';
                Expression = { $Mailbox.DisplayName }
            },
            @{
                Label = 'UserPrincipalName';
                Expression = { $Mailbox.UserPrincipalName }
            },
            @{
                Label = 'PrimarySmtpAddress';
                Expression = { $Mailbox.PrimarySmtpAddress }
            },
            @{
                Label = 'MailboxType';
                Expression = { $Mailbox.Recipienttypedetails }
            },
            @{
                Label = 'MailboxCreatedon';
                Expression = { $Mailbox.WhenMailboxCreated }
            },
            @{
                Label = 'Lastloggedon';
                Expression = { $Inactivemailbox.LastLogonTime }
            },
            @{
                Label = 'Inactive';
                Expression = { $DaysInactive }
            }
    )
    $Output += $Instance
}
$Output | Export-Csv $Outpath -Encoding UTF8 -NoTypeInformation