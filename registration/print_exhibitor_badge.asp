<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim sBadgeString
	dim sReceiptString
	dim lngRecs
	dim strSQL
	dim conDB
	dim rs
	dim sPayType
	dim sAmount
	dim sProgramAmount
	dim sName
	dim sCompany
	dim sCompany1
	dim sCompany2
	dim s
	dim t
	dim i
	dim intPrintBadge
	dim intPrintBadgeReceipt
	dim intPrintProgramReceipt
	dim intProgramAmount
	dim intTotalAmount
	dim bolPrintHeader
	dim badge_string_num
	dim programtest
	dim programtester

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
			Set objFS = CreateObject("Scripting.FileSystemObject")
		Set objPrint = objFS.CreateTextFile("LPT1:", True)
		objPrint.Write("^XA^A0N,30,30^FO100,50^FD" & Session("PRINT_ME") & "^FS^XZ")
		objPrint.Close
	Session("PRINT_ME") = " " & Session("PRINT_ME") & ","
	Session("NEEDSRECEIPT") = " " & Session("NEEDSRECEIPT") & ","
	Session("PRINT_PROGRAM_RECEIPT") = " " & Session("PRINT_PROGRAM_RECEIPT") & ","
	Session("PROGRAM_COUNT") = Session("PROGRAM_COUNT") & ","
	
	programtest = split(Session("PROGRAM_COUNT"),",")

	strSQL = "select e.company_name, eg.first_name, eg.last_name, eg.guest_id, e.preregistered, b.badge_num " & _
			 "from exhibitor e, exhibitor_guest eg, badge b " & _
			 "where e.exhibitor_id = eg.exhibitor_id " & _
			 "and eg.guest_id = b.guest_id " & _
			 "and eg.exhibitor_id = " & Session("EXHIBITOR_ID") & _
			 " and eg.status = 'A' "
			 
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)

	Do While Not rs.EOF
		sAmount = "0.00"
		
		badge_string_num = " " & rs("badge_num") & ","
		
		intPrintBadge = instr(Session("PRINT_ME"), badge_string_num)
		
		intPrintBadgeReceipt = instr(Session("NEEDSRECEIPT"), badge_string_num)
		
		intPrintProgramReceipt = instr(Session("PRINT_PROGRAM_RECEIPT"), badge_string_num)
		
		sName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 27)
		
		sCompany =  rs("company_name")
		
		If Len(sCompany) > 25 Then
'find the closest space before character 26, get it's number
			s = Mid(sCompany, 1, 26)  's = United States Air Force Ba
			
			if StrComp(Mid(s, 26, 1), Chr(32)) = 0 then
				For i = 26 To 1 Step -1
					t = StrComp(Mid(s, i, 1), Chr(32))
					
					If t = 0 Then
						sCompany1 = Mid(sCompany, 1, i)
						
						sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
						
						Exit For
					End If
				Next
			else
				For i = 31 To 1 Step -1
					t = StrComp(Mid(s, i, 1), Chr(32))
					
					If t = 0 Then
						sCompany1 = Mid(sCompany, 1, i)
						
						sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
						
						Exit For
					End If
				Next
			end if
'place all text before it in sCompany1, and all after it in sCompany2
		Else
			sCompany1 = sCompany
			
			sCompany2 = "&nbsp"
		End If
		
		If rs("preregistered") = "Y" Then
			sPayType = "Preregistration"
		Else
			sPayType = "Registration"
		End If
		
		If instr(Session("NEEDSRECEIPT"), rs("badge_num")) Then
			sAmount = "100.00"
		Else
			sAmount = " 0.00"
		End If

		If intPrintBadge > 0 Then

			sBadgeString = "^XA" & _
"^A0N,100,100^FO0,70^FB816,1,,C^FD" & sName & "^FS" & _
"^A0N,60,60^FO0,200^FB816,1,,C^FD" & sCompany1 & " " & sCompany2 & "^FS" & _
"^A0N,40,40^FO0,300^FB816,1,,C^FDExhibitor^FS" & _
"^XZ"			

		
		End If
		
		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 then
			sReceiptString= "^XA" & _
"^A0N,40,40^FO40,200^REGISTRATION DETAILS:^FS" & _
"^A0N,40,40^FO40,250^FDPAY TYPE^FS" & _
"^A0N,40,40^FO350,250^FDDATE^FS" & _
"^A0N,40,40^FO650,250^FDAMOUNT^FS"
		End If
			
		If intPrintBadgeReceipt > 0 then
			sReceiptString = sReceiptString & "^A0N,40,40^FO320,320^FD" & Date & "^FS" & _
							"^A0N,40,40^FO650,320^FD" & sAmount & "^FS" 
		End If
			
		If intPrintProgramReceipt > 0 Then
			For i = 1 to 7
				programtester = instr(Left(programtest(i),1),rs("badge_num") - 1)
				
				If programtester > 0 then
					sProgramAmount = Right(programtest(i),Len(programtest(i)) - 2)
					
					sProgramAmount = 20 * sProgramAmount
										
					Exit For
				End If
			Next			
		End If
			
		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 Then
			sReceiptString = sReceiptString & "^XZ"				
			
			Set objFS = CreateObject("Scripting.FileSystemObject")
			Set objPrint = objFS.CreateTextFile("LPT1:", True)
			objPrint.Write(Session("PRINT_ME"))
			objPrint.Close
		End If

		
		rs.movenext		
	Loop
	
	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing		


%>
<html>
	<head>
	</head>
	<body>
	</body>
</html>
