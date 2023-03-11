Import-Module ActiveDirectory
CompDN = "(&(objectCategory=computer)(objectClass=computer)(cn=$env:COMPUTERNAME))"
$strCompDN = [string]([adsisearcher]$CompDN).FindOne().Properties.distinguishedname
$objComp = [ADSI]("LDAP://"+$strCompDN)

# quit if computer is a server or DC
if (($strCompDN -like '*Controller*') -or ($strCompDN -like '*SERVER*')) { exit }

$strUsrLogin = $env:ctobin
$strDateStamp = Get-Date -f 'yyyy-MM-dd@HH:mm'
$IPPattern = "^\d+\.\d+\.\d+\.\d+$"

$colNICs = gwmi Win32_NetworkAdapterConfiguration
if ($colNICs.Count -gt 0) {
foreach ($objNIC in $colNICs){
        if ($objNIC.DefaultIPGateway) {
            $arrIP = $objNIC.IPAddress
            for ($i=0; $i -lt $colNICs.Count; $i++) { 
            if ($arrIP[$i] -match $IPPattern) { $strIP = $arrIP[$i]; $strMAC = $objNIC.MACAddress }
            }
        }
    }
}

$objComp.Description = $strDateStamp + " - " + $strUsrLogin + " - " + $strIP
$objComp.extensionAttribute1 = $strUsrLogin
$objComp.extensionAttribute2 = $strIP
$objComp.extensionAttribute3 = $strMAC
$objComp.SetInfo()
