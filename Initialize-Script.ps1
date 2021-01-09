param (
    [String]$inputPath = "$env:homeshare\VDI-UserData\Download\generic\inputs\",
    [String]$fileName = "mailbox_list.csv",
    [String]$user = $null
)

if ($user) {
    $allMailboxes = @($user)
} else {
    $allMailboxes = Get-Content "$inputPath\$fileName"
}