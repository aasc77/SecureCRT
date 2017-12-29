Function SendCommands(header,csv)
STARTLOG csv
Dim Timer, Hostname
Timer = 60 

crt.Screen.Synchronous = 	false
crt.Screen.Synchronous =	True

crt.Screen.Send ("term length 0" & vbCR )		
crt.Screen.Send ("sh run" & vbCR )	

crt.Screen.Send ("show cdp neighbor detail" & vbCR )	


crt.Screen.Send ("sh run | inc hostname" & vbCR )
Hostname = crt.Screen.WaitForStrings("hostname 00","hostname 0","hostname 1",Timer)

IF Hostname > 1 THEN

crt.Screen.Send ("sh ver" & vbCR )
stroutput = crt.Screen.WaitForStrings("CISC01","CISC02",Timer)	



IF stroutput = 1 THEN

crt.Screen.Send ("config t" & vbCR ) 

crt.Screen.Send ("ip dhcp excluded-address "& csv(4) & vbCR )'csv(4)= Default Router IP

crt.Screen.Send ("no interface GigabitEthernet0/0.70" & vbCR )

crt.Screen.Send ("ip dhcp pool 1"& vbCR )
crt.Screen.Send ("network " & csv(3) & " 255.255.255.252" & vbCR )'csv(3)= network IP
crt.Screen.Send ("default-router " & csv(4) & vbCR ) 'csv(4)= Default Router IP
 )

crt.Screen.Send ("end"& vbCR )

END IF
END IF
crt.Screen.Send ("exit"& vbCR ) 
crt.Screen.WaitForString("closed.")



 
 
 
STOPLOG
SendCommands = stroutput &"-"& Hostname
  
End Function





