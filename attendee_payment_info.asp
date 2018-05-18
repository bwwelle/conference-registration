<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sFirst
	dim sLast
	dim sCity
	dim sState
	dim sATTENDEE_LIST
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsAttendee
	dim iRegID
	dim rsPaymentAct
	dim sPayType
	dim arAttendeeField
	
	sATTENDEE_LIST = Session("ATTENDEE_LIST")
	
	If sATTENDEE_LIST = "" Then
		Response.Write("No Payment Activity for this Attendee.")
	Else
'parse for first and last name, city and state
		arAttendeeField = split(sATTENDEE_LIST, " | ")
		
		sLast = arAttendeeField(0)
		
		sFirst = arAttendeeField(1)
		
		sCity = arAttendeeField(2)
		
		sState = arAttendeeField(3)
		
		If InStr(1, sState, "^") Then
			sState = Trim(Left(sState, Len(sState)-1))
		End If
	
		Set conDB = Server.CreateObject("ADODB.Connection")
		
		conDB.Open strCon	
		
'use names to get registration_id
		strSQL = "SELECT registration_id FROM registration "
		
		strSQL = strSQL & "WHERE first_name = '" & ManageQuotes(Trim(sFirst)) & "' "
		
		strSQL = strSQL & "AND last_name = '" & ManageQuotes(Trim(sLast)) & "' "			
		
		strSQL = strSQL & "AND city = '" & ManageQuotes(Trim(sCity)) & "' "
		
		strSQL = strSQL & "AND state = '" & ManageQuotes(Trim(sState)) & "' "
		
		Set rsAttendee = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		If lngRecs = -1 Then
			iRegID = rsAttendee("registration_id")
		Else
'log it somewhere
		End If
		
		Set rsAttendee = Nothing
		
'use registration_id to get payment info
		strSQL = "SELECT r.first_name, r.last_name, "
		
		strSQL = strSQL & "isnull(l.amount_paid, 0) amount_paid, isnull(l.payment_type, '') payment_type, l.acttimestamp, isnull(l.complimentary, 'N') complimentary, l.tot_amount_due, isnull(l.outstanding, '') outstanding, isnull(l.initials, '') initials "
		
		strSQL = strSQL & "FROM log_badge_payment_activity l, registration r "
		
		strSQL = strSQL & "WHERE r.registration_id = l.registration_id "
		
		strSQL = strSQL & "AND r.registration_id = " & iRegID
		
		strSQL = strSQL & " ORDER BY l.acttimestamp desc"
		
		Set rsPaymentAct = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		Response.Write("<table border=0 cellspacing=""0""><tr bgcolor=""#AAD5FF""><td width=150><div align=""left""><strong>Attendee Name</strong></div></td>")
		
		Response.Write("<td width=120><div align=""right""><strong>Total Amount Due</strong></div></td>")
		
		Response.Write("<td width=100><div align=""right""><strong>Amount Paid</strong></div></td>")
		
		Response.Write("<td width=120><div align=""center""><strong>Complimentary</strong></div></td>")
		
		Response.Write("<td width=120><div align=""left""><strong>Payment Type</strong></div></td>")
		
		Response.Write("<td width=100><div align=""left""><strong>Initials</strong></div></td>")
		
		Response.Write("<td width=200><div align=""right""><strong>Date and Time</strong></div></td></tr>")
		
		Do While Not rsPaymentAct.EOF
			Response.Write("<tr><td width=150><div align=""left"">" & rsPaymentAct("last_name") & ", " & rsPaymentAct("first_name") & "</div></td>")
			
			Response.Write("<td width=120><div align=""right"">" & rsPaymentAct("tot_amount_due") & "</div></td>")
			
			Response.Write("<td width=100><div align=""right"">" & rsPaymentAct("amount_paid") & "</div></td>")
			
			Response.Write("<td width=120><div align=""center"">" & rsPaymentAct("complimentary") & "</div></td>")
			
			sPayType = ""
			
			If instr(rsPaymentAct("payment_type"), "CC") Then
				sPayType = " Credit "
			End If
			
			If instr(rsPaymentAct("payment_type"), "CH") Then
				sPayType = sPayType & " Check "
			End If
			
			If instr(rsPaymentAct("payment_type"), "CA") Then
				sPayType = sPayType & " Cash/Trav "
			End If
			
			If instr(rsPaymentAct("payment_type"), "NC") Then
				sPayType = sPayType & " NoCharge "
			End If
			
			Response.Write("<td width=120><div align=""left"">" & sPayType & "</div></td>")
			
			Response.Write("<td width=100><div align=""left"">" & rsPaymentAct("initials") & "</div></td>")
			
			Response.Write("<td width=200><div align=""right"">" & rsPaymentAct("acttimestamp") & "</div></td></tr>")
			
			rsPaymentAct.Movenext
		Loop
		
		Set rsPaymentAct = Nothing
		
		conDB.Close
		
		Set conDB = Nothing
		
		Response.Write("<tr><td colspan=7><div align=""center""><a href=""registration.asp"">Return to Registration</a></div></td></tr>")
		
		Response.Write("</table>")	
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>The Registration Payment Information</title>
		<style type="text/css">		
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>	
	<body></body>
</html>
