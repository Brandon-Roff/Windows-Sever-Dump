Get-ADUser -filter * -searchbase 'OU="Staff",OU=Users,DC=Doamain,DC=internal' | Set-ADUser -Replace @{Atteibutename='Staff';'}
