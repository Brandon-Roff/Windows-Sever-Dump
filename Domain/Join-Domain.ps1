$domain = "PADC"
$password = "BradminPutin" | ConvertTo-SecureString -asPlainText -Force
$username = "$domain\bradmin" 
$credential = New-Object System.Management.Automation.PSCredential($username,$password)
Add-Computer -DomainName $domain -Credential $credential