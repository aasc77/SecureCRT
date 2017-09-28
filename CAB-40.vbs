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
stroutput = crt.Screen.WaitForStrings("CISCO29","CISCO89",Timer)	' checks for showrooms



IF stroutput = 1 THEN

crt.Screen.Send ("config t" & vbCR ) 

crt.Screen.Send ("ip dhcp excluded-address "& csv(4) & vbCR )'csv(4)= Default Router IP

crt.Screen.Send ("no interface GigabitEthernet0/0.70" & vbCR )
crt.Screen.Send ("interface GigabitEthernet0/0.70" & vbCR )
crt.Screen.Send ("description DIGITAL-COMMUNITY-MIRROR [/m]" & vbCR )
crt.Screen.Send ("encapsulation dot1Q 70" & vbCR )
crt.Screen.Send ("ip address " & csv(4) & " 255.255.255.252" & vbCR )
crt.Screen.Send ("no ip redirects" & vbCR )
crt.Screen.Send ("no ip proxy-arp"& vbCR )
crt.Screen.Send ("ip nbar protocol-discovery"& vbCR )
crt.Screen.Send ("zone-member security vend"& vbCR )
crt.Screen.Send ("service-policy input STORE-QOS-IN"& vbCR )

crt.Screen.Send ("ip dhcp pool DIGITAL-COMMUNITY-MIRROR"& vbCR )
crt.Screen.Send ("network " & csv(3) & " 255.255.255.252" & vbCR )'csv(3)= network IP
crt.Screen.Send ("default-router " & csv(4) & vbCR ) 'csv(4)= Default Router IP
crt.Screen.Send ("domain-name store.net" & vbCR )
crt.Screen.Send ("dns-server 12.130.160.50 8.8.8.8"& vbCR )

crt.Screen.Send ("end"& vbCR )

END IF
END IF
crt.Screen.Send ("exit"& vbCR ) 
crt.Screen.WaitForString("closed.")



 
 
 
STOPLOG
SendCommands = stroutput &"-"& Hostname
  
End Function





