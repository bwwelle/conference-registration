<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
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
	
	If Request.Form("report_exhibitor") <> "" Then
		Session("SORT_ORDER_EXHIBITOR") = Request.Form("SORT_BY")
	Else
		Session("SORT_ORDER_EXHIBITOR") = "company_name"
	End if
	If Session("SORT_FIELD")= "" then
		Session("SORT_FIELD") = "badge_count"
	else
		Session("SORT_FIELD") = Request.Form("SORT_BY")
	End if 	
	
	If Session("SORT_DIRECTION") = "" Then
		Session("SORT_DIRECTION") = "asc"
	Elseif Session("SORT_DIRECTION") = "asc" Then
		Session("SORT_DIRECTION") = "desc"
	Elseif Session("SORT_DIRECTION") = "desc" Then
		Session("SORT_DIRECTION") = "asc"
	End If
	
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
		
	'get the counts and compay names for all those exhibitors who have picked up at least one badge
	strSQL = "select count(distinct b.badge_id) badge_count, e.company_name company_name "
	strSQL = strSQL & "from badge b, log_badge_activity l, exhibitor_guest eg, exhibitor e "
	strSQL = strSQL & "where b.badge_id = l.badge_id and b.guest_id = eg.guest_id and "
	strSQL = strSQL & "e.exhibitor_id = eg.exhibitor_id and b.badge_type = 'E'  group by company_name "
	strSQL = strSQL & "UNION select '0' badge_count, e.company_name company_name from badge b, exhibitor_guest eg, exhibitor e "
	strSQL = strSQL & "where b.guest_id = eg.guest_id and e.exhibitor_id = eg.exhibitor_id and b.badge_type = 'E' and b.badge_id not "
	strSQL = strSQL & "in (select badge_id from log_badge_activity) and e.company_name not in(select e.company_name company_name from badge b"
	strSQL = strSQL & ", log_badge_activity l, exhibitor_guest eg, exhibitor e where b.badge_id = l.badge_id and b.guest_id = eg.guest_id and "
	strSQL = strSQL & "e.exhibitor_id = eg.exhibitor_id and b.badge_type = 'E'  group by company_name having count(distinct b.badge_id) >0"
	strSQL = strSQL & ") order by " & Session("SORT_FIELD") & " " & Session("SORT_DIRECTION")
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sTableString = "<table border='0' width='400' cellspacing='0'><form name='report_exhibitor' method=post action='report_exhibitors_present.asp'>"
	sTableString = sTableString & "<tr bgcolor='#AAD5FF'><td width='100' align='right'><a href='report_select.asp'>Back To Reports</a></td></tr>"
	sTableString = sTableString & "<tr bgcolor='#AAD5FF'>"
	sTableString = sTableString & "<td width='300 align='left'><input type='submit' name='SORT' value='Sort' >"
	sTableString = sTableString & "<input type='radio' name='SORT_BY' value='company_name'>Exhibitors</strong></td>"
	sTableString = sTableString & "<td width='100' align='right'><input type='radio' name='SORT_BY' value='badge_count'>Badges</strong></td></tr>"
	iColorCount = 0

	Do While Not rs.EOF
		iColorCount = iColorCount + 1
		If (iColorCount mod 2) Then
			sColorCode = "white"
		Else
			sColorCode = "#DFEFFF"
		End If
		
		sTableString = sTableString & "<tr bgcolor='" & sColorCode & "'><td width='300' align='left'>" & rs("company_name") & "</td>" 
		sTableString = sTableString & "<td width='100' align=""right"">" & rs("badge_count") & "</td></tr>"
		
		rs.MoveNext	
	Loop
	sTableString = sTableString & "</form></table>"
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
		<%=sTableString%>
	</body>
</html>
