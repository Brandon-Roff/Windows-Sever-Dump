SET /P PCNAME=Please enter your name: 
REG ADD HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName /v ComputerName /t REG_SZ /d %PCNAME% /f