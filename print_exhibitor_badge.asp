<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim sBadgeString
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
	
	Session("PRINT_ME") = " " & Session("PRINT_ME") & ","
	Session("NEEDSRECEIPT") = " " & Session("NEEDSRECEIPT") & ","
	Session("PRINT_PROGRAM_RECEIPT") = " " & Session("PRINT_PROGRAM_RECEIPT") & ","
	Session("PROGRAM_COUNT") = Session("PROGRAM_COUNT") & ","
	
	programtest = split(Session("PROGRAM_COUNT"),",")

	strSQL = "select e.company_name, eg.first_name, eg.last_name, eg.guest_id, 			e.preregistered, b.badge_num " & _
			 "from exhibitor e, exhibitor_guest eg, badge b " & _
			 "where e.exhibitor_id = eg.exhibitor_id " & _
			 "and eg.guest_id = b.guest_id " & _
			 "and eg.exhibitor_id = " & Session("EXHIBITOR_ID") & _
			 " and eg.status = 'A' "
			 
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
		
	Do While Not rs.EOF
		sAmount = "0.00"
		
		badge_string_num = " " & rs("badge_num") & ","
		
		intPrintBadge = instr(Session("PRINT_ME"), badge_string_num)
		
		intPrintBadgeReceipt = instr(Session("NEEDSRECEIPT"), badge_string_num)
		
		intPrintProgramReceipt = instr(Session("PRINT_PROGRAM_RECEIPT"), badge_string_num)
		
		sName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 27)
		
		sCompany =  rs("company_name")
		
		If Len(sCompany) > 30 Then
'find the closest space before character 26, get it's number
			s = Mid(sCompany, 1, 31)  's = United States Air Force Ba
			
			For i = 31 To 1 Step -1
				t = StrComp(Mid(s, i, 1), Chr(32))
				
				If t = 0 Then
					sCompany1 = Mid(sCompany, 1, i)
					
					sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
					
					Exit For
				End If
			Next
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
			sAmount = "135.00"
		Else
			sAmount = " 0.00"
		End If
		
		If intPrintBadge > 0 Then			
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:36px'><b>" & sName & "</b></td></tr>" 
			
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:10px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany1 & "</b></td></tr>" 
			
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany2 & "</b></td></tr>" 
			
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:8px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:16px'><b>" & rs("guest_id") & "</b></td>" 
			
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:20px'><b>EXHIBITOR</b></td></tr>"
		End If
		
		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 then
			sBadgeString = sBadgeString & "<tr><td width='110' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>"
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>"
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='left' style='font-size:16px'><b>REGISTRATION  PAYMENT DETAILS:</b></td></tr>" 
			
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>PAYMENT TYPE</b></td>" 
			
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>DATE</b></td>"
			
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>"
		End If
			
		If intPrintBadgeReceipt > 0 then
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>" & sPayType & "</b></td>"
			
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>" & Date & "</b></td>"
			
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>$" & sAmount & "</b></td></tr>"
			
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>"
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
			
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>Program(s)</b></td>"
			
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>" & Date & "</b></td>" 
			
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>$" & sProgramAmount & "</b></td></tr>"
		End If
			
		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 Then
			sBadgeString = sBadgeString & "<tr><td colspan='2' align='left' style='font-size:16px'><b>" & sName & "</b></td>" 
			
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:16px'><b>" & rs("guest_id") & "</b></td></tr>" 
			
			sBadgeString = sBadgeString & "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" 
			
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:16px'><b>$" & sProgramAmount + sAmount & "</b></td></tr>" 
		End If
		
		rs.movenext		
	Loop
	
	sBadgeString = sBadgeString & "</table>"
	
	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing
%>
<html>
	<head>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body onLoad="javascript:window.print();javascript:window.close();">
		<form name="printform">
			<% =sBadgeString %>
		</form>
	</body>
</html>
