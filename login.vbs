# $language = "VBScript"
# $interface = "1.0"
' Author: Angel Serrano



Sub Include(file)
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(file, 1)
  str = f.ReadAll
  f.Close
  ExecuteGlobal str
End Sub

Function WriteLog(LoginData,n)

LoginResult = Array("login successful","login failed","Name or service not known","Connection refused","Connection closed by remote host","TIMEOUT")
LoginData(4).writeline ( LoginData(0) &","& LoginResult(n) &","& LoginData(1) &","& LoginData(2) &","& "Time:" &","& LoginData(3)&","& LoginData(5))
WriteLog = output

End Function

Function Continue(strMsg, strTitle)

nButtons = vbYesNo + vbDefaultButton2

nIcon = vbQuestion

If MsgBox(strMsg, nButtons + nIcon, strTitle) <> vbYes Then
Continue = False
Else
Continue = True
End If
End Function

Function Time

time = month(Now())&"/"&Day(Now())&"/"& Year(Now())&"/"&"-"& Hour(Now())&":" &Minute(Now()) &":" &Second(Now())
Time = time
End Function

SUB STARTLOG(csv)
 
	ip = csv(2)
	crt.Session.Log False
	srlogfile = ip&".conf"
	crt.Session.LogFileName = srlogfile
	crt.Session.Log True

End SUB

SUB STOPLOG
crt.Session.Log False
End SUB

SUB ExitDevice(csv)

crt.Screen.Send Chr(26)  
crt.Screen.Send Chr(26)
crt.Screen.Send (vbCR & vbCR )

STOPLOG
END SUB  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

SUB Main
	Const ForReading = 1
	Const ForWriting = 2
  Dim fso,fso2,ip, login2,login3, asses, Timer, protocol,tab,strHideUsername,strHidePassword,n, result
  tab = chr(9)
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
		'strCSV		= crt.Dialog.Prompt("Enter File Name:", "Input", "", False)
		Batch 		= crt.Dialog.Prompt("Enter Batch Number:", "batch", "", False)
		strUN 		= crt.Dialog.Prompt("Enter Username:", "Login", "", False)
		strPW 		= crt.Dialog.Prompt("Enter password:", "Login", "", True)
		'strCount 	= crt.Dialog.Prompt("Devices to run before prompt:", "Prompt interval", "", False)
		
	
	strUsername = strUN
	strPassword = strPW
	'strInputFile  =  strCSV
	
	'strUsername = ""
	'strPassword = ""
	strInputFile = "input.csv"
	
	strPasscode =""
	strEnablePassword =""
	Timer = 30
	Timer2 = 15 
	
    protocol = 0 'telnet = 1 ;SSH = 0  
	
	
	
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set fso  = CreateObject("Scripting.FileSystemObject")
		Set fso2 = CreateObject("Scripting.FileSystemObject")
	
		Set AllDevices 		=   fso2.OpenTextFile(month(Now())&"_"&Day(Now())&"-"& Year(Now())&"-"& Hour(Now())&"_" &Minute(Now()) &"-" &Second(Now())& "_output-report.csv", ForWriting, True)
		Set input  			=   fso.OpenTextFile(strInputFile,   ForReading, False)
	
  crt.Screen.Synchronous = True
  crt.Screen.Send ( vbCR & vbCR )
  
  Dim count 
  count = 0
  
  
  
  Do While input.AtEndOfStream <> True
	
	
	
    data= input.Readline
	
	If instr (data, "IP") = 0 then   '1
	
	
	
	Dim csv
	csv = Split(data,",")
	
	IF csv(1) = Batch THEN   '2
	
	ip  = csv(2)
	LoginData  = Array(ip,strHideUsername,strHidePassword,Time,AllDevices,result)
	IF protocol = 1 THEN
	crt.Screen.Send ("telnet "& ip  & vbCR )
	END IF 

	IF protocol = 0 THEN


	crt.Screen.Send ( vbCR & vbCR ) 
	crt.Screen.WaitForStrings  ("$")
	Include csv(0)
	crt.Screen.Send ("ssh -oStrictHostKeyChecking=no -oUserKnownHostsFile=/dev/null "& strUsername &"@" & ip  & vbCR )
	count = count + 1

	END IF
		
	login = crt.Screen.WaitForStrings("asscode:","assword:","sername","Name or service not known","Connection refused","Connection closed by remote host",Timer)
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 	
	IF login = 0 THEN
	WriteLog LoginData, 5
	ExitDevice(csv)
	END IF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
	

'PC_TART
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
	IF login = 1 THEN 
	crt.Screen.Send (strPasscode & vbCR )

					login2 = crt.Screen.WaitForStrings("#","asscode",Timer)
					IF login2 = 0 THEN 
					WriteLog LoginData ,5 
					ExitDevice(csv)
					END IF
					IF login2 = 1 THEN
					
					LoginData(5) = SendCommands(header,csv)
					WriteLog LoginData, 0
					
					END IF
					IF login2 = 2 THEN 
					WriteLog LoginData, 1 
					ExitDevice(csv)
					END IF
	END IF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
'PC_END
	

	
'PW_START
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
	IF login = 2 THEN
	crt.Screen.Send (strPassword & vbCR )

					login2 = crt.Screen.WaitForStrings  ("#","assword",">",Timer)
					IF login2 = 0 THEN
					WriteLog LoginData, 5
					ExitDevice(csv)
					END IF
					IF login2 = 1 THEN
					
					LoginData(5) = SendCommands(header,csv)
					WriteLog LoginData, 0
					
					END IF
					IF login2 = 2 THEN
					
					WriteLog LoginData, 1
					ExitDevice(csv)
					
					END IF
					IF login2 = 3 THEN
					crt.Screen.Send (vbCR )
					crt.Screen.Send ("en" & vbCR )
					crt.Screen.Send (strEnablePassword & vbCR )

								login3 = crt.Screen.WaitForStrings  ("#","assword",Timer)
								IF login3= 0 THEN
								WriteLog LoginData, 5 
								ExitDevice(csv)
								END IF
								IF login3 = 1 THEN
								
						LoginData(5) = SendCommands(header,csv)
						WriteLog LoginData, 0
									
								END IF
								IF login3 = 2 THEN
								
								WriteLog LoginData, 1  
								ExitDevice(csv)
								
								END IF
					
					END IF
					
	END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PW_END



'UN_PW_EN_START
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF login = 3 THEN
	crt.Screen.Send (strUsername & vbCR )
asses = crt.Screen.WaitForStrings ("assword","asscode")

IF asses = 1 THEN	
	crt.Screen.Send (strPassword & vbCR )

					login2 = crt.Screen.WaitForStrings  ("#","assword",">",Timer)
					IF login2 = 0 THEN
					WriteLog LoginData, 5
					ExitDevice(csv)
					END IF
					IF login2 = 1 THEN
					
					h
					WriteLog LoginData, 0
					
					END IF
					IF login2 = 2 THEN
					WriteLog LoginData, 1 
					ExitDevice(csv)
					END IF
					IF login2 = 3 THEN
					crt.Screen.Send (vbCR )
					crt.Screen.Send ("en" & vbCR )
					crt.Screen.Send (strEnablePassword & vbCR )

								login3 = crt.Screen.WaitForStrings  ("#","assword",Timer)
								IF login3 = 0 THEN
								WriteLog LoginData, 1 
								ExitDevice(csv)
								END IF
								IF login3 = 1 THEN
								
					LoginData(5) = SendCommands(header,csv)
					WriteLog LoginData, 0
 					
					END IF
								IF login3 = 2 THEN
								WriteLog LoginData, 1  
								ExitDevice(csv)
								END IF
	
					END IF
	END IF				
END IF
IF asses = 2 THEN
crt.Screen.Send (strPasscode & vbCR )
					login2 = crt.Screen.WaitForStrings  ("#","asscode",Timer)
					IF login2 = 0 THEN
					WriteLog LoginData, 5 
					ExitDevice(csv)
					END IF
					IF login2 = 1 THEN
					
					LoginData(5) = SendCommands(header,csv)
					WriteLog LoginData, 0
					END IF
END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
'UN_PW_EN_END
	

'Failures_START
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	

	IF login = 4 THEN
	WriteLog LoginData, 2 
	ExitDevice(csv)
	END IF
	IF login = 5 THEN
	WriteLog LoginData, 3 
	ExitDevice(csv)
	END IF
	IF login = 6 THEN
	WriteLog LoginData, 4 
	ExitDevice(csv)
	END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	 
'Failures_END	
	else 
	
	Dim header
	header = Split(data,",")
	
	'MsgBox header(0)& header(0)
	Dim var 
	Set var = CreateObject("Scripting.Dictionary")
 
	'var.Add header(0), 60
	'result.Add "Name", "Tony"

	
	'MsgBox var(header(0))

	End if '1
END IF '2
Loop
	 
crt.Screen.Synchronous = False

End Sub



