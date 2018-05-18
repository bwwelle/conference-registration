<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sEXHIBITOR_ID
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsPaymentAct
	dim sPayType
	
	sEXHIBITOR_ID = Session("EXHIBITOR_ID")
	
	If sEXHIBITOR_ID = "" Then
		Response.Write("No Booth Payment Activity for this Exhibitor.")
	Else
	
		Set conDB = Server.CreateObject("ADODB.Connection")
		
		conDB.Open strCon	
			
'use registration_id to get payment info
		strSQL = "SELECT e.company_name, "
		
		strSQL = strSQL & "isnull(l.amount_paid, 0) amount_paid, isnull(l.payment_type, '') payment_type, l.acttimestamp, l.tot_amount_due, isnull(l.initials, '') initials "
		
		strSQL = strSQL & "FROM log_booth_payment_activity l, exhibitor e "
		
		strSQL = strSQL & "WHERE e.exhibitor_id = l.exhibitor_id "
		
		strSQL = strSQL & "AND e.exhibitor_id = " & sEXHIBITOR_ID
		
		strSQL = strSQL & " ORDER BY l.acttimestamp desc"
		
		Set rsPaymentAct = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		Response.Write("<table border=0 cellspacing=""0""><tr bgcolor=""#AAD5FF""><td width=150><div align=""left""><strong>Exhibitor Name</strong></div></td>")
		
		Response.Write("<td width=120><div align=""right""><strong>Total Amount Due</strong></div></td>")
		
		Response.Write("<td width=100><div align=""right""><strong>Amount Paid</strong></div></td>")
		
		Response.Write("<td width=120><div align=""left""><strong>Payment Type</strong></div></td>")
		
		Response.Write("<td width=100><div align=""left""><strong>Initials</strong></div></td>")
		
		Response.Write("<td width=200><div align=""right""><strong>Date and Time</strong></div></td></tr>")
		
		Do While Not rsPaymentAct.EOF
			Response.Write("<tr><td width=150><div align=""left"">" & rsPaymentAct("company_name") & "</div></td>")
			
			Response.Write("<td width=120><div align=""right"">" & rsPaymentAct("tot_amount_due") & "</div></td>")
			
			Response.Write("<td width=100><div align=""right"">" & rsPaymentAct("amount_paid") & "</div></td>")
			
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
			
			Response.Write("<td width=120><div align=""left"">" & sPayType & "</div></td>")
			
			Response.Write("<td width=100><div align=""left"">" & rsPaymentAct("initials") & "</div></td>")
			
			Response.Write("<td width=200><div align=""right"">" & rsPaymentAct("acttimestamp") & "</div></td></tr>")
			
			rsPaymentAct.Movenext
		Loop
		
		Set rsPaymentAct = Nothing
		
		conDB.Close
		
		Set conDB = Nothing
		
		Response.Write("<tr><td colspan=7><div align=""center""><a href=""exhibitor.asp"">Return to Select Exhibitor</a></div></td></tr>")
		
		Response.Write("</table>")	
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>The Registration Exhibitor Booth Payment Information</title>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body>
<!-- BEGIN CONTENT -->
	</body>
</html>
