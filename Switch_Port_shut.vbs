Function SendCommands(header,csv)
STARTLOG csv

crt.Screen.Synchronous = 	false
crt.Screen.Synchronous =	True
  '' uses csv as input to determin which ports to shut.  checks ports before making change
  ''''''change below only
stroutput = "Completed"
crt.Screen.Send ("term length 0"& vbCR )		
crt.Screen.Send ("sh run"& vbCR )  
crt.Screen.Send ("config t"& vbCR ) 

For n = 2 to 50
strport = csv(n)
Select Case strport

Case "1"

crt.Screen.Send ("do show int status | inc Fa0/"&n-2&vbCR )
result = crt.Screen.WaitForString ("notconnect",3)
 'MsgBox result
IF result = -1 THEN
crt.Screen.Send ("int Fa0/"&n-2&vbCR )
crt.Screen.Send ("shut"& vbCR )
	END IF
IF result = 0 THEN
stroutput = "One or ports were in use, not shut"
	END IF
	
End Select 

Next

crt.Screen.Send ("end"& vbCR )  
crt.Screen.Send ("exit"& vbCR ) 
login = crt.Screen.WaitForString("closed.")

''''''change above only

STOPLOG
SendCommands = stroutput
  
End Function





