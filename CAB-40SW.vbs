Function SendCommands(header,csv)
STARTLOG csv

crt.Screen.Synchronous = 	false
crt.Screen.Synchronous =	True

crt.Screen.Send ("term length 0"& vbCR )		
crt.Screen.Send ("sh run" & vbCR )	
crt.Screen.Send ("sh ver" & vbCR )	

crt.Screen.Send ("config t" & vbCR ) 

crt.Screen.Send ("vlan 70"& vbCR )
crt.Screen.Send ("name DIGITAL-COMMUNITY-MIRROR" & vbCR )

crt.Screen.Send ("interface FastEthernet0/30" & vbCR )
crt.Screen.Send ("description DIGITAL-COMMUNITY-MIRROR [/m]" & vbCR )
crt.Screen.Send ("no shut" & vbCR )
crt.Screen.Send ("switchport access vlan 70" & vbCR )
crt.Screen.Send ("switchport mode access" & vbCR )
crt.Screen.Send ("spanning-tree portfast" & vbCR )
crt.Screen.Send ("end" & vbCR )
crt.Screen.Send ("wr" & vbCR )
crt.Screen.Send (vbCR )
crt.Screen.Send ("exit"& vbCR ) 
crt.Screen.WaitForString("closed.")



stroutput = "Completed"
 
 
 
STOPLOG
SendCommands = stroutput
  
End Function





