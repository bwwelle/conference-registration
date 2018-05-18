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
	dim sEXHIBITOR_ID
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsAttendee
	dim iRegID
	dim rsBadgeAct
	dim arAttendeeField
	
	sEXHIBITOR_ID = Session("EXHIBITOR_ID")
	
	If sEXHIBITOR_ID = "" Then
		Response.Write("No badges picked up for this exhibitor.")
	Else
		Set conDB = Server.CreateObject("ADODB.Connection")
		
		conDB.Open strCon	
		
'get badge info
		strSQL ="SELECT eg.last_name + ', ' + eg.first_name as lfname, b.badge_num, l.pickup_name, l.acttimestamp "
		
		strSQL = strSQL & "FROM log_badge_activity l, badge b, exhibitor_guest eg "
		
		strSQL = strSQL & "WHERE l.badge_id = b.badge_id "
		
		strSQL = strSQL & "AND b.guest_id = eg.guest_id "
		
		strSQL = strSQL & "AND eg.exhibitor_id = " & sEXHIBITOR_ID
		
		strSQL = strSQL & " ORDER BY lfname"
		
		Set rsBadgeAct = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
		Response.Write("<table border=0 cellspacing=""0""><tr bgcolor=""#AAD5FF""><td width=150><div align=""left""><strong>Guest Name</strong></div></td>")
		
		Response.Write("<td width=60><div align=""center""><strong>Badge #</strong></div></td>")
		
		Response.Write("<td width=150><div align=""center""><strong>Badge Picked Up By</strong></div></td>")
		
		Response.Write("<td width=200><div align=""right""><strong>Date and Time</strong></div></td></tr>")
		
		Do While Not rsBadgeAct.EOF
			Response.Write("<tr><td width=150><div align=""left"">" & rsBadgeAct("lfname") & "</div></td>")
			
			Response.Write("<td width=60><div align=""center"">" & rsBadgeAct("badge_num") & "</div></td>")
			
			Response.Write("<td width=150><div align=""left"">" & rsBadgeAct("pickup_name") & "</div></td>")
			
			Response.Write("<td width=200><div align=""right"">" & rsBadgeAct("acttimestamp") & "</div></td></tr>")
			
			rsBadgeAct.Movenext
		Loop
		
		Set rsBadgeID = Nothing
		
		conDB.Close
		
		Set conDB = Nothing
		
		Response.Write("<tr><td colspan=3><div align=""center""><a href=""exhibitor.asp"">Return to Registration</a></div></td></tr>")
		
		Response.Write("</table>")	
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>The Registration Badge Information</title>
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
