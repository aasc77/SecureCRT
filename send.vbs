
Function SendCommands(header,csv)
STARTLOG csv

crt.Screen.Synchronous = 	false
crt.Screen.Synchronous =	True
  
  ''''''change below only
crt.Screen.Send ("term length 0"& vbCR )		
crt.Screen.Send ("sh run"& vbCR )  
crt.Screen.Send ("exit"& vbCR )  
login = crt.Screen.WaitForString("closed.")
''''''change above only
STOPLOG
SendCommands = output
  
End Function





