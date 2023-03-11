Set objNetwork = CreateObject("WScript.Network")

objNetwork.AddWindowsPrinterConnection "\\OPA-SRV16-PR01.padc.internal\Follow Me_Mono"
objNetwork.AddWindowsPrinterConnection "\\OPA-SRV16-PR01.padc.internal\Follow Me color"

ObjNetwork.SetDefaultPrinter "\\OPA-SRV16-PR01.padc.internal\Follow Me_Mono"