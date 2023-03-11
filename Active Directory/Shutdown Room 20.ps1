
$R20 = 4285 

DO

{

 "Updating Group Policy on OPA-R20-$R20"

Stop-Computer -ComputerName OPA-R20-$R20 -Force -Credential PADC\bradmin GreatIris2022@

 "Successfully Updated Group Policy on OPA-R20-$R20"
   ""
     
 $R20++



} While ($R20 -le 4315)