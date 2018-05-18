<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="/includes/dbfunctions.asp" -->
<!--#include virtual="/includes/dbglobals.asp" -->
<!--#include virtual="/includes/adovbs.inc"-->

<%
	dim conDB
	dim strSQL
	dim lngRecs
	dim rs
	dim iColorCount
	dim sColorCode
	dim sTableString
	dim sTableString2
	dim sSTATE
	dim sSortby

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
	
	If Request.Form("SORT_BY") <> "" Then
		sSortby = Request.Form("SORT_BY")
	Else
		sSortby = "last_name, first_name"
	End if

'get the counts and compay names for all those exhibiotrs who have picked up at least one badge
	strSQL = "select l.acttimestamp, r.last_name, r.first_name, b.badge_type,r.preregistered, l.tot_amount_due, l.amount_paid, l.payment_type "
	
	strSQL = strSQL & "from log_badge_payment_activity l, registration r, badge b "
	
	strSQL = strSQL & "where l.registration_id = r.registration_id "
	
	strSQL = strSQL & "and b.registration_id = r.registration_id "
	
	strSQL = strSQL & "and r.registration_id <> '' "
	
	strSQL = strSQL & "order by " & sSortby 
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sTableString = "<table border='0' width='100%' cellspacing='0'>" & _
				   "<tr><td colspan='9' align='center'><a href='report_select.asp'>Return to Reports</a></td></tr>" & _
				   "<tr bgcolor='#AAD5FF'>" & _
				   "<td width='20%' align='center'><strong><input type='radio' name='SORT_BY' value='last_name, first_name'>Name</strong></td> " & _
				   "<td width='15%' align='right'><strong><input type='radio' name='SORT_BY' value='acttimestamp'>Date</strong></td>" & _
				   "<td width='15%' align='right'><strong><input type='radio' name='SORT_BY' value='payment_type, last_name'>Payment Type</strong></td>" & _
				   "<td colspan='6' align='left'><input type='submit' name='SORT' value='Sort' ></td>" & _
	
				   "<tr bgcolor='#AAD5FF'><td colspan='9' align='center'><b>Attendee Payment Log</b></td></tr>" & _
				   "<tr bgcolor='#AAD5FF'>" & _
				   "<td width='20%' align='left'><b>Time Logged</b></td>" & _
				   "<td width='15%' align='left'><b>Last Name</b></td>" & _
				   "<td width='15%' align='left'><b>First Name</b></td>" & _
				   "<td width='5%' align='center'><b>Badge Type</b></td>" & _
				   "<td width='5%' align='center'><b>Prereg?</b></td>" & _
				   "<td width='10%' align='right'><b>Total Amount Due</b></td>" & _
				   "<td width='10%' align='right'><b>Amount Paid</b></td>" & _
				   "<td width='10%' align='center'><b>Payment Type</b></td></tr>"
				   
	iColorCount = 0
	
	Do While Not rs.EOF
		iColorCount = iColorCount + 1
		
		If (iColorCount mod 2) Then
			sColorCode = "white"
		Else
			sColorCode = "#DFEFFF"
		End If
		
		sTableString = sTableString & "<tr bgcolor='" & sColorCode & "'>" & _
									   "<td width='20%' align='left'><b>" & rs("acttimestamp") & "</b></td>" & _
									   "<td width='15%' align='left'><b>" & rs("last_name") & "</b></td>" & _
									   "<td width='15%' align='left'><b>" & rs("first_name") & "</b></td>" & _
									   "<td width='5%' align='center'><b>" & rs("badge_type") & "</b></td>" & _
									   "<td width='5%' align='center'><b>" & rs("preregistered") & "</b></td>" & _
									   "<td width='10%' align='right'><b>" & rs("tot_amount_due") & "</b></td>" & _
									   "<td width='10%' align='right'><b>" & rs("amount_paid") & "</b></td>" & _
									   "<td width='10%' align='center'><b>" & rs("payment_type") & "</b></td></tr>"
									   
		rs.MoveNext	
	Loop
	
	sTableString = sTableString & "<tr><td colspan='9' align='center'><a href='report_select.asp'>Return to Reports</a></td></tr></table>"
	
	Set rs = Nothing
		
	conDB.Close
	
	Set conDB = Nothing
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration - Report View</title>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body>
		<form name="report_attendee_payment" method=post action="report_attendee_payment_log.asp">
			<%=sTableString%>
		</form>
	</body>
</html>
