Dim g_shell
Set g_shell = CreateObject("WSCript.Shell")

Dim g_objExcel
Set g_objExcel = Nothing

Dim g_strMyDocs, g_strWkBkPath
'g_strMyDocs = g_shell.SpecialFolders("MyDocuments")
'g_strWkBkPath = g_strMyDocs & "\MyExcelData.xls"
'g_strWkBkPath = "C:\Book2.xlsx"

g_strWkBkPath = "C:\Lulu-Lemon-CPE_Config_Generator_Tool_v4f2.xlsm"
' g_strError is a global variable that is used to store error messages
' within called functions so that the error messages are available to
' other (calling) functions or subroutines.
Dim g_strError

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Main()

step34 = "PASS"
step35 = "PASS"
step36 = "PASS"
step37 = "PASS"	
step38 = "PASS"
step39 = "PASS"	



    Dim strSearchFor, strSearchCol
    strSearchFor = "Site ID"
    
   
    strSearchCol = "A"
    
    Do
        ' Prompt user for data to find
        strSearchFor = crt.Dialog.Prompt(_
            "Please Enter Site ID", _
            "Search Spreadsheet For...", _
            strSearchFor)
        
        ' Check for "cancel"
        If strSearchFor = "" Then Exit Do
        
        Dim nRowFound, vData
       
        nRow = Lookup(g_strWkBkPath, strSearchFor, strSearchCol, vRowData, "T")
        crt.Dialog.MessageBox "Account found in row " &nRow  
      
        If nRow = 0 Then
            crt.Dialog.MessageBox _
                """" & strSearchFor & """ was not found in column " & _
                """" & strSearchCol & """ in the specified spreadsheet."
        Else
      
SSData2 = DoWorkWithRowData(vRowData)

'crt.Dialog.MessageBox "this is out " &SSData2(3)
        
        End If
        
    Loop
    
    ' Close Excel (only if it was ever opened)
    If Not g_objExcel Is Nothing Then g_objExcel.Quit
    
	
	Dim Timer	
	strAudit = "PASSED"
Timer = 5

Audit_Time = month(Now())&"_"&Day(Now())&"_"& Year(Now())&"_"& Hour(Now())&"-" &Minute(Now()) &"-" &Second(Now())
	crt.Session.Log False
    srlogfile = SSData2(0)&".Audit-"& Audit_Time &".log"
	crt.Session.LogFileName = srlogfile
	crt.Session.Log True

crt.Screen.Synchronous = False
crt.Screen.Synchronous = True

crt.screen.Send ("term length 0" & vbCR )
crt.screen.Send ("sh run " & vbCR )

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~STEP 34 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
crt.Screen.Send ("sh ver | inc FastEthernet interfaces"& vbCR )  
ports = crt.Screen.WaitForStrings("24 ", "48 ")
 
IF ports = 1 THEN
p = 24
END IF 
IF ports = 2 THEN
p = 47
END IF 

For n = 1 to p

crt.Screen.Send ("show int status | inc Fa0/"&n&" "&vbCR )
result = crt.Screen.WaitForStrings ("connected","notconnect","disabled",60)

Select Case n

Case 25,26,27,45,46,47

IF result <> 1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n& vbCrLf & "Interface should be connected-Step 34"
step34 ="FAILED"

END IF

Case 1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,29,30 	
IF result <> 2 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n& vbCrLf &"Interface should be notconnect-Step 34"
step34 ="FAILED"
END IF

Case 31,32,33,34,35,36,48 	
IF result <> 3 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n& vbCrLf &" Interface should be disabled-Step 34"
step34 ="FAILED"
END IF

End Select 

Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

crt.Screen.Send (vbCR ) 
crt.screen.Send ("sh int GigabitEthernet0/1"& vbCR )
IntG1 = crt.Screen.WaitForString("administratively down",Timer)
IF IntG1 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Interface Audit Failed for GigabitEthernet0/1"& vbCrLf & "Interface should be Admin Down-Step 34" 
step34 ="FAILED"
END IF

crt.Screen.Send (vbCR )
crt.screen.Send ("sh int GigabitEthernet0/2"& vbCR )
IntG2 = crt.Screen.WaitForString("administratively down",Timer)
IF IntG2 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Interface Audit Failed for GigabitEthernet0/2"& vbCrLf & "Interface should be Admin Down-Step 34" 
step34 ="FAILED"
END IF
crt.Screen.Send (vbCR )
crt.screen.Send ("sh int GigabitEthernet0/4"& vbCR )
IntG4 = crt.Screen.WaitForString("administratively down",Timer)
IF IntG4 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Interface Audit Failed for GigabitEthernet0/4"& vbCrLf & "Interface should be Admin Down-Step 34" 
step34 ="FAILED"
END IF

crt.Screen.Send (vbCR )
crt.screen.Send ("show int vlan238"& vbCR )
IntVlan_128 = crt.Screen.WaitForString("is up, line protocol is up",Timer)
IF IntVlan_128 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Interface Audit Failed for vlan238"& vbCrLf & "Interface should be up up-Step 34" 
step34 ="FAILED"
END IF

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~CHECK DESCRIPTION Fa0/x ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
crt.Screen.Send ("sh ver | inc FastEthernet interfaces"& vbCR )  
ports = crt.Screen.WaitForStrings("24 ", "48 ")
 
IF ports = 1 THEN
p = 24
END IF 
IF ports = 2 THEN
p = 47
END IF 

For n = 1 to p

crt.Screen.Send ("show int status | inc Fa0/"&n&" "&vbCR )
desc = crt.Screen.WaitForStrings ("POS server interfa","Traffic-Device","BACK OFFICE","Security camera","WAP","Unassigned","E-Media device int","Store printer inte","UPS","PDU","RFID Reader","LR01 - G0/0 (LAN t","AT&T Netgate - LAN","LR01 - G0/2 (WAN I","AT&T MRS Router","Connection to vide","DIGITAL-COMMUNITY",Timer)

Select Case n

Case 1,2,3,4,5,6
IF desc <> 1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be -POS server interfa-STEP 39"
step39 ="FAILED"
END IF

Case 7,8,9
IF desc <> 2 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be -TRAFFIC-DEVICE-STEP 39"
step39 ="FAILED"
END IF

Case 10,11,12
IF desc <> 3 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be -BACK OFFICE-STEP 39"
step39 ="FAILED"
END IF

Case 13,14,15,16,17,18,19,20,21,22,23,24 
IF desc <> 4 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be -Security camera in-STEP 39"
step39 ="FAILED"
END IF

Case 25,26,27,28,29
IF desc <> 5 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - WAP-STEP 39"
step39 ="FAILED"
END IF


Case 30
IF desc <> 17 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - DIGITAL-COMMUNITY-STEP 39"
step39 ="FAILED"
END IF


'Case 31,32,33,34,35,36
'IF desc <> 6 THEN 
'strAudit = "FAILED"
'crt.Dialog.MessageBox "Audit Failed on Fa0/"&n&"Description should be - Unassigned-STEP 39"
'step39 ="FAILED"
'END IF

Case 37 
IF desc <> 7 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - E-Media device int-STEP 39"
step39 ="FAILED"
END IF

Case 38 
IF desc <> 8 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - Store printer inte-STEP 39"
step39 ="FAILED"
END IF

Case 39 
IF desc <> 9 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - UPS-STEP 39"
step39 ="FAILED"
END IF

Case 40 
IF desc <> 10 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - PDU-STEP 39"
step39 ="FAILED"
END IF

Case 41,42,43,44 
IF desc <> 11 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - RFID Reader-STEP 39"
step39 ="FAILED"
END IF

Case 45 
IF desc <> 12 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - LR01 - G0/0 (LAN t-STEP 39"
step39 ="FAILED"
END IF

Case 46 
IF desc <> 13 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - AT&T Netgate - LAN-STEP 39"
step39 ="FAILED"
END IF

Case 47 
IF desc <> 14 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - LR01 - G0/2 (WAN I-STEP 39"
step39 ="FAILED"
END IF

Case 48 
IF desc <> 15 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Audit Failed on Fa0/"&n & vbCrLf &"Description should be - AT&T MRS Router [/-STEP 39"
step39 ="FAILED"
END IF

End Select 

Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~STEP 35 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
crt.Screen.Send (vbCR )
crt.screen.Send ("show cdp neighbors | inc AP01"& vbCR )
AP1 = crt.Screen.WaitForString("Fas 0/25",Timer)
IF AP1 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "AP1 Audit Failed"& vbCrLf &"AP1 not found on Fas 0/25-STEP 35" 
step35 ="FAILED"
END IF
crt.Screen.Send (vbCR )
crt.screen.Send ("show cdp neighbors | inc AP02"& vbCR )
AP2 = crt.Screen.WaitForString("Fas 0/26",Timer)
IF AP2 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "AP2 Audit Failed"& vbCrLf &"AP2 not found on Fas 0/26-STEP 35" 
step35 ="FAILED"
END IF
crt.Screen.Send (vbCR )
crt.screen.Send ("show cdp neighbors | inc AP03"& vbCR )
AP3 = crt.Screen.WaitForString("Fas 0/27",Timer)
IF AP3 <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "AP3 Audit Failed"& vbCrLf &"AP3 not found on Fas 0/27-STEP 35" 
step35 ="FAILED"
END IF

crt.screen.Send ("show run | count permit"& vbCR )
Permit = crt.Screen.WaitForString("regexp = 20",Timer)
IF Permit <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Unexpected Permit count"& vbCrLf &"should be 20-STEP 36" 
step36 ="FAILED"
END IF

crt.screen.Send ("show run | count remark"& vbCR )
Remark = crt.Screen.WaitForString("regexp = 10",Timer)
IF Remark <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Unexpected Permit count"& vbCrLf &"should be 10-STEP 37" 
step37 ="FAILED"
END IF

crt.screen.Send ("show run | count access-list"& vbCR )
ACL = crt.Screen.WaitForString("regexp = 14",Timer)
IF ACL <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Unexpected Permit count"& vbCrLf &"should be 14-STEP 38" 
step38 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 10   enet"& vbCR )
vlan = crt.Screen.WaitForString("10   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 10 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 62   enet"& vbCR )
vlan = crt.Screen.WaitForString("62   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 62 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 66   enet"& vbCR )
vlan = crt.Screen.WaitForString("66   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 66 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 78   enet"& vbCR )
vlan = crt.Screen.WaitForString("78   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 78 not found-STEP 39"
step39 ="FAILED" 
END IF

crt.screen.Send ("show vlan | inc 86   enet"& vbCR )
vlan = crt.Screen.WaitForString("86   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 86 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 110   enet"& vbCR )
vlan = crt.Screen.WaitForString("110   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 110 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 142   enet"& vbCR )
vlan = crt.Screen.WaitForString("142   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 142 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 158   enet"& vbCR )
vlan = crt.Screen.WaitForString("158   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 158 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 190   enet"& vbCR )
vlan = crt.Screen.WaitForString("190  enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 190 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 191   enet"& vbCR )
vlan = crt.Screen.WaitForString("191   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 191 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 206   enet"& vbCR )
vlan = crt.Screen.WaitForString("206   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 206 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 222   enet"& vbCR )
vlan = crt.Screen.WaitForString("222   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 222 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 238   enet"& vbCR )
vlan = crt.Screen.WaitForString("238   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 238 not found-STEP 39" 
step39 ="FAILED"
END IF

crt.screen.Send ("show vlan | inc 246   enet"& vbCR )
vlan = crt.Screen.WaitForString("246   enet",Timer)
IF vlan <> -1 THEN 
strAudit = "FAILED"
crt.Dialog.MessageBox "Vlan 246 not found-STEP 39" 
step39 ="FAILED"
END IF

'crt.Dialog.MessageBox "Audit for "& SSData2(0)& " has " & strAudit	

crt.Dialog.MessageBox "Audit for "& SSData2(0) & vbCrLf	& "Step 34: "&step34 & vbCrLf & "Step 35: "&step35 & vbCrLf & "Step 36: "&step36 & vbCrLf	& "Step 37: "&step37 & vbCrLf	& "Step 38: "&step38 & vbCrLf	& "Step 39: "& step39	




crt.Session.Log False	

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function  DoWorkWithRowData(vRowData)

	strNumber = vRowData(Col("A"))
	Str3RDOctect = vRowData(Col("B"))
	strSubnet = vRowData(Col("C"))
    strVoiceSubnet = vRowData(Col("D"))
    strHostname = vRowData(Col("E"))
	strSwitchHostname = vRowData(Col("F"))
	strAtt_AS = vRowData(Col("G"))
	strLulu_AS = vRowData(Col("H"))
	strLocation = vRowData(Col("I"))
	strAVPN_Link_IP = vRowData(Col("N"))
	strAVPN_CER =  vRowData(Col("O"))
	strAVPN_PER = vRowData(Col("P"))
	
	
	crt.Dialog.MessageBox _
                 " Working Account # " & strNumber & " located in " & strLocation & " with hostname " & strHostname 
	
	
	
	'crt.Dialog.MessageBox strNumber 
	'crt.Dialog.MessageBox strSubnet
	'crt.Dialog.MessageBox strHostname
	'crt.Dialog.MessageBox strSwitchHostname
	'crt.Dialog.MessageBox strAtt_AS
	'crt.Dialog.MessageBox strLulu_AS
	'crt.Dialog.MessageBox strAVPN_IP
	'crt.Dialog.MessageBox strAVPN_CER

	

  SSData  = Array(strNumber, Str3RDOctect, strSubnet, strVoiceSubnet, strHostname, strSwitchHostname, strAtt_AS, strLulu_AS, strLocation, strAVPN_Link_IP, strAVPN_CER, strAVPN_PER)
  'crt.Dialog.MessageBox "this is " &SSData(10)
  
  DoWorkWithRowData =  SSData

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Lookup(strWkBkPath, _
                strSearchFor, _
                strColumnToSearch, _
                ByRef vRowDataArray, _
                nLastColData)
   
    If g_objExcel Is Nothing Then
        On Error Resume Next
        Set g_objExcel = CreateObject("Excel.Application")
        nError = Err.Number
        strErr = Err.Description
        On Error Goto 0
        
        If nError <> 0 Then
            crt.Dialog.Prompt _
                "Error: " &  strErr & vbcrlf & _
                vbcrlf 
                
            Exit Function
        End If
    End If
    
    ' Now load the workbook (Read-Only), and get a reference to the first sheet
    On Error Resume Next
    Set objWkBk = g_objExcel.Workbooks.Open(strWkBkPath, 0, True)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    
    If nError <> 0 Then
        g_objExcel.Quit
        crt.Dialog.MessageBox _
            "Error loading spreadsheet """ & strWkBkPath & """:" & vbcrlf & _
            vbcrlf & _
            strErr
        Exit Function
    End If
    
    Set objSheet = objWkBk.Sheets(2)
    
    ' Look in specified column for the given data:
    Set objSearchRange = objSheet.Columns(strColumnToSearch)
    Set objFoundRange = Nothing
    Set objFoundRange = objSearchRange.Find(strSearchFor)
    If Not objFoundRange Is Nothing Then
        Lookup = objFoundRange.Row
        ReDim vRowDataArray(Col(nLastColData))
        For nColIndex = 1 To Col(nLastColData)
            vRowDataArray(nColIndex) = objFoundRange.Cells(1, nColIndex).Value
        Next
    Else
        Lookup = 0
        Exit Function
    End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Col(strLetterRef)

    Dim nColumnValue
    nColumnValue = 0


    For nLetterIndex = Len(strLetterRef) To 1 Step -1

        nMultiplier = Len(strLetterRef) - nLetterIndex

        nCurLetter = ASC(UCase(Mid(strLetterRef, nLetterIndex, 1))) - 64
        nColumnValue = nColumnValue + (nCurLetter * (26 ^ nMultiplier))
    Next
    
    Col = nColumnValue
End Function



