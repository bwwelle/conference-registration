<% @Language=VBScript %>
<% Response.Buffer="true"%>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim conDB
	dim strSQL
	dim lngRecs
	dim rs
	dim sREGDATE
	dim sSearchDateReg
	dim sSearchDateAct
	dim sSearchDateAct2
	dim sPrereg
	dim iCountExhibitor
	dim iCountAdult
	dim iCountGuest
	dim iCountCommercial
	dim iCountCollege
	dim iCountClinician
	dim iCountStaff
	dim iSumBadgeCA
	dim iSumBadgeCC
	dim iSumBadgeCH
	dim iSumBoothCA
	dim iSumBoothCC
	dim iSumBoothCH
	dim iTotalCountBadges
	dim iTotalBadges
	dim iTotalBooths
	dim iTotalCash
	dim iTotalCheck
	dim iTotalCredit
	dim iTotal

	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
	sPrereg = "Y"
	
	'get the counts for all the separate badge types printed for the selected date
	'for exhibitor
	strSQL = "select count(distinct b.badge_id) badges "
	strSQL = strSQL & "from  badge b, exhibitor_guest eg, exhibitor e "
	strSQL = strSQL & "where eg.guest_id = b.guest_id " 
	strSQL = strSQL & "and e.exhibitor_id = eg.exhibitor_id "
	strSQL = strSQL & "and eg.status = 'A' " 
	strSQL = strSQL & "and b.badge_type = 'E' "
	strSQL = strSQL & "and e.preregistered = 'Y' "		
	'debug 
	'response.Write("strSQL = " + strSQL + "  ")
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCountExhibitor = rs("badges")
	Set rs = Nothing
	
	'for adult/director
	strSQL = "select count(distinct b.badge_id) badges "
	strSQL = strSQL & "from badge b, registration r "
	strSQL = strSQL & "where b.registration_id = r.registration_id "
	strSQL = strSQL & "and b.badge_type = 'D' "
	strSQL = strSQL & "and r.preregistered = 'Y'"
	'debug 
	'response.Write("strSQL = " + strSQL + "  ")	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCountAdult = rs("badges")
	Set rs = Nothing

	iCountGuest = 0
	iCountCommercial = 0
	iCountCollege = 0
	iCountClinician = 0
	iCountStaff = 0
	
	'now get the badge payments for that day
	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
			' "where payment_type = 'CA' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBadgeCA = 0  'rs("amount_paid")
	'Set rs = Nothing

	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
			' "where payment_type = 'CC' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBadgeCC = 0  'rs("amount_paid")
	'Set rs = Nothing

	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
			' "where payment_type = 'CH' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBadgeCH = 0  'rs("amount_paid")
	'Set rs = Nothing
	
	'and the booth payents for that day
	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
			' "where payment_type = 'CA' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBoothCA = 0  'rs("amount_paid")
	'Set rs = Nothing

	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
			' "where payment_type = 'CC' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBoothCC = 0  'rs("amount_paid")
	'Set rs = Nothing

	'strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
			' "where payment_type = 'CH' and acttimestamp " & sSearchDateAct
	'Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iSumBoothCH = 0  'rs("amount_paid")
	'Set rs = Nothing
	

	iTotalCountBadges = iCountExhibitor + iCountAdult + iCountGuest + iCountCommercial + iCountCollege + iCountClinician + iCountStaff
	iTotalBooths = iSumBoothCA + iSumBoothCC + iSumBoothCH
	iTotalBadges = iSumBadgeCA + iSumBadgeCC + iSumBadgeCH
	iTotalCash =  iSumBoothCA + iSumBadgeCA
	iTotalCheck = iSumBoothCH + iSumBadgeCH
	iTotalCredit = iSumBoothCC + iSumBadgeCC
	iTotal = iTotalBooths + iTotalBadges
	
	conDB.Close
	Set conDB = Nothing


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration - Report Preregistered</title>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body>
		<table border="0" width="500" cellspacing="0">
			<tr bgcolor="#AAD5FF">
				<td colspan="6" align="center">
					<b>
						Preregistered Attendees
					</b>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td width="250" align="left">
					<b>
						Badge Type
					</b>
				</td>
				<td width="250" align="right">
					<b>
						Badge Count
					</b>
				</td>
			</tr>
			<tr>
				<td width="250" align="left">
					Exhibitor
				</td>
				<td width="250" align="right">
					<%=iCountExhibitor%>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td width="250" align="left">
					Adult/Director
				</td>
				<td width="250" align="right">
					<%=iCountAdult%>
				</td>
			</tr>
			<tr>
				<td width="250" align="left">
					Guest
				</td>
				<td width="250" align="right">
					<%=iCountGuest%>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td width="250" align="left">
					Commercial
				</td>
				<td width="250" align="right">
					<%=iCountCommercial%>
				</td>
			</tr>
			<tr>
				<td width="250" align="left">
					College Student
				</td>
				<td width="250" align="right">
					<%=iCountCollege%>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td width="250" align="left">
					Clinician
				</td>
				<td width="250" align="right">
					<%=iCountClinician%>
				</td>
			</tr>
			<tr>
				<td width="250" align="left">
					Staff
				</td>
				<td width="250" align="right">
					<%=iCountStaff%>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td width="250" align="left" >
					<b>
						Total Badges
					</b>
				</td>
				<td width="250" align="right">
					<b>
						<%=iTotalCountBadges%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="6" align="center">
					<a href="report_select.asp">
						Return to Select Reports
					</a>
				</td>
			</tr>
		</table>
	</body>
</html>
