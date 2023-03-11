Import-Module ActiveDirectory
$SMTPServer = "in-v3.mailjet.com"
$SMTPPort = "587"
$From = "asset.ormistonpark.org.uk"
$To = "broff@ormistonpark.org.uk"
$Subject = "User Locked Out"
$SMTPUsername = "c6abea4ae4bc4265130284f08d2ca77f"
$SMTPPassword = "13784320d861d8e8a7687f21f31bea70"
$SMTPCredentials = New-Object System.Management.Automation.PSCredential($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force))


$LockedOutUsers = Search-ADAccount -LockedOut

If ($LockedOutUsers) {
    ForEach ($LockedOutUser in $LockedOutUsers) {
        $Body = "User $($LockedOutUser.Name) has been locked out of Active Directory."
        Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $From -To $To -Subject $Subject -Body $Body -Credential $SMTPCredentials
    }
}