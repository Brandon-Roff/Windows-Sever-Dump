$R20 = 4285 

DO

{

 "Updating Group Policy on OPA-R20-$R20"

Invoke-GPUpdate -Computer OPA-R20-$R20

 "Successfully Updated Group Policy on OPA-R20-$R20"
   ""
     
 $R20++



} While ($R20 -le 4304)