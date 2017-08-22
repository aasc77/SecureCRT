Function SendCommands(header,csv)
STARTLOG csv

crt.Screen.Synchronous = 	false
crt.Screen.Synchronous =	True
''''''change below only

stroutput = "Completed"
crt.Screen.Send ("term length 0"& vbCR )		
crt.Screen.Send ("sh run"& vbCR )  
crt.Screen.Send ("sh ver | inc FastEthernet interfaces"& vbCR )  
ports = crt.Screen.WaitForStrings("24 ", "48 ")

'MsgBox ports

IF ports = 1 THEN
p = 24
END IF 
IF ports = 2 THEN
p = 48
END IF 

'MsgBox p
FOR n = 1 TO p
strport = csv(n)
SELECT CASE strport

CASE "1"

crt.Screen.Send ("config t"& vbCR ) 
crt.Screen.Send ("do show int status | inc Fa0/"&n&" "&vbCR )
result = crt.Screen.WaitForStrings ("notconnect","connected","disabled",3)
'MsgBox result
IF result = 1 THEN
crt.Screen.Send ("int Fa0/"&n&vbCR )
crt.Screen.Send ("shut"& vbCR )
END IF
IF result = 0 THEN
stroutput = "One or ports were in use and not shut"
END IF
	
END SELECT 

NEXT

crt.Screen.Send ("end"& vbCR )  
crt.Screen.Send ("exit"& vbCR ) 
crt.Screen.Send ("#wr"& vbCR ) 
login = crt.Screen.WaitForString("closed.")

''''''change above only

STOPLOG
SendCommands = stroutput
  
End Function





