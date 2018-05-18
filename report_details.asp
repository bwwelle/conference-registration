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

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
		
'get the counts for all the director badge types for the week to date, by state
	strSQL = "select count(distinct b.badge_id) badge_count, r.state "
	
	strSQL = strSQL & "from registration r, badge b, log_badge_activity l "
	
	strSQL = strSQL & "where r.registration_id = b.registration_id "
	
	strSQL = strSQL & "and l.acttimestamp >= 'Dec 12 2005' "
	
	strSQL = strSQL & "and l.acttimestamp <= 'Dec 17 2005' "
	
	strSQL = strSQL & "and b.badge_id = l.badge_id "
	
	strSQL = strSQL & "and b.badge_type = 'D' "
	
	strSQL = strSQL & "and r.state <> '' "
	
	strSQL = strSQL & "group by r.state "
	
	strSQL = strSQL & "order by r.state "
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	iCount = 0
	
	sTableString = "<table border='0' width='400' cellspacing='0'>" & _
				   "<tr><td colspan='2' align='center'><a href='report_select.asp'>Return to Select Reports</a></td></tr>" & _
				   "<tr bgcolor='#AAD5FF'><td colspan='2' align='center'><b>Week To Date Badge Details By State</b></td></tr>" & _
				   "<tr bgcolor='#AAD5FF'><td width='300' align='left'><b>States</b></td>" & _
				   "<td width='100' align='right'><b>Badge Counts</b></td></tr>"
				   
	iColorCount = 0
	
	Do While Not rs.EOF
		iColorCount = iColorCount + 1
		
		If (iColorCount mod 2) Then
			sColorCode = "white"
		Else
			sColorCode = "#DFEFFF"
		End If
		
		sSTATE = rs("state")
		
		If sSTATE = "AL" Then sSTATE = "Alabama"
		If sSTATE = "AK" Then sSTATE = "Alaska"
		If sSTATE = "AZ" Then sSTATE = "Arizona"
		If sSTATE = "AR" Then sSTATE = "Arkansas"
		If sSTATE = "CA" Then sSTATE = "California"
		If sSTATE = "CO" Then sSTATE = "Colorado"
		If sSTATE = "CT" Then sSTATE = "Connecticut"
		If sSTATE = "DE" Then sSTATE = "Delaware"
		If sSTATE = "DC" Then sSTATE = "District of Columbia"
		If sSTATE = "FL" Then sSTATE = "Florida"
		If sSTATE = "GA" Then sSTATE = "Georgia"
		If sSTATE = "GU" Then sSTATE = "Guam"
		If sSTATE = "HI" Then sSTATE = "Hawaii"
		If sSTATE = "ID" Then sSTATE = "Idaho"
		If sSTATE = "IL" Then sSTATE = "Illinois"
		If sSTATE = "IN" Then sSTATE = "Indiana"
		If sSTATE = "IA" Then sSTATE = "Iowa"
		If sSTATE = "KS" Then sSTATE = "Kansas"
		If sSTATE = "KY" Then sSTATE = "Kentucky"
		If sSTATE = "LA" Then sSTATE = "Louisiana"
		If sSTATE = "ME" Then sSTATE = "Maine"
		If sSTATE = "MD" Then sSTATE = "Maryland"
		If sSTATE = "MA" Then sSTATE = "Massachusetts"
		If sSTATE = "MI" Then sSTATE = "Michigan"
		If sSTATE = "MN" Then sSTATE = "Minnesota"
		If sSTATE = "MS" Then sSTATE = "Mississippi"
		If sSTATE = "MO" Then sSTATE = "Missouri"
		If sSTATE = "MT" Then sSTATE = "Montana"
		If sSTATE = "NE" Then sSTATE = "Nebraska"
		If sSTATE = "NV" Then sSTATE = "Nevada"
		If sSTATE = "NH" Then sSTATE = "New Hampshire"
		If sSTATE = "NJ" Then sSTATE = "New Jersey"
		If sSTATE = "NM" Then sSTATE = "New Mexico"
		If sSTATE = "NY" Then sSTATE = "New York"
		If sSTATE = "NC" Then sSTATE = "North Carolina"
		If sSTATE = "ND" Then sSTATE = "North Dakota"
		If sSTATE = "OH" Then sSTATE = "Ohio"
		If sSTATE = "OK" Then sSTATE = "Oklahoma"
		If sSTATE = "OR" Then sSTATE = "Oregon"
		If sSTATE = "PA" Then sSTATE = "Pennsylvania"
		If sSTATE = "PR" Then sSTATE = "Puerto Rico"
		If sSTATE = "RI" Then sSTATE = "Rhode Island"
		If sSTATE = "SC" Then sSTATE = "South Carolina"
		If sSTATE = "SD" Then sSTATE = "South Dakota"
		If sSTATE = "TN" Then sSTATE = "Tennessee"
		If sSTATE = "TX" Then sSTATE = "Texas"
		If sSTATE = "UT" Then sSTATE = "Utah"
		If sSTATE = "VT" Then sSTATE = "Vermont"
		If sSTATE = "VI" Then sSTATE = "Virgin Islands"
		If sSTATE = "VA" Then sSTATE = "Virginia"
		If sSTATE = "WA" Then sSTATE = "Washington"
		If sSTATE = "WV" Then sSTATE = "West"
		If sSTATE = "WI" Then sSTATE = "Wisconsin"
		If sSTATE = "WY" Then sSTATE = "Wyoming"
		If sSTATE = "AB" Then sSTATE = "Alberta"
		If sSTATE = "BC" Then sSTATE = "British"
		If sSTATE = "MB" Then sSTATE = "Manitoba"
		If sSTATE = "NB" Then sSTATE = "New Brunswick"
		If sSTATE = "NS" Then sSTATE = "New South Wales"
		If sSTATE = "NL" Then sSTATE = "Newfoundland"
		If sSTATE = "NT" Then sSTATE = "Northwest Territories"
		If sSTATE = "NS" Then sSTATE = "Nova Scotia"
		If sSTATE = "NU" Then sSTATE = "Nunavut"
		If sSTATE = "ON" Then sSTATE = "Ontario"
		If sSTATE = "PE" Then sSTATE = "Prince Edward Island"
		If sSTATE = "QC" Then sSTATE = "Quebec"
		If sSTATE = "QLD" Then sSTATE = "Queensland"
		If sSTATE = "SK" Then sSTATE = "Saskatchewan"
		If sSTATE = "SA" Then sSTATE = "South Australia"
		If sSTATE = "TAS" Then sSTATE = "Tasmania"
		If sSTATE = "VIC" Then sSTATE = "Victoria"
		If sSTATE = "W AUST" Then sSTATE = "Western Australia"
		If sSTATE = "YT" Then sSTATE = "Yukon Territory"
		If sSTATE = "AE" Then sSTATE = "APO AE"
		If sSTATE = "NT" Then sSTATE = "Northern Territory"
		If sSTATE = "AC" Then sSTATE = "Australian Capital Territory"	
		
		sTableString = sTableString & "<tr bgcolor='" & sColorCode & "'><td width='300' align='left'>" & rs("state") & "</td>" & _
									  "<td width='100' align=""right"">" & rs("badge_count") & "</td></tr>"
									  
		iCount = iCount + rs("badge_count")
		
		rs.MoveNext	
	Loop
	
	sTableString = sTableString & "<tr bgcolor='#AAD5FF'><td width='300' align='left'><b>Total</b></td>" & _
								  "<td width='100' align=""right"">" & iCount & _
								  "</td></tr><tr><td colspan='2' align='center'><a href='report_select.asp'>Return to Select Reports</a></td></tr></table>"
								  
	Set rs = Nothing
		
'get the counts for all the separate badge types for the week to date, by country
	strSQL = "select count(distinct b.badge_id) badge_count, r.country "
	
	strSQL = strSQL & "from registration r, badge b, log_badge_activity l "
	
	strSQL = strSQL & "where r.registration_id = b.registration_id "
	
	strSQL = strSQL & "and l.acttimestamp <= 'Dec 17 2005' "
	strSQL = strSQL & "and b.badge_id = l.badge_id "
	
	strSQL = strSQL & "and b.badge_type = 'D' "
	
	strSQL = strSQL & "group by r.country "
	
	strSQL = strSQL & "order by r.country "
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	iCount2 = 0
	
	sTableString2 = "<table border='0' width='400' cellspacing='0'>" & _
				   "<tr bgcolor='#AAD5FF'><td colspan='2' align='center'><b>Week To Date Badge Details By Country</b></td></tr>" & _
				   "<tr bgcolor='#AAD5FF'><td width='300' align='left'><b>Countries</b></td>" & _
				   "<td width='100' align='right'><b>Badge Counts</b></td></tr>"
	iColorCount = 0
	
	Do While Not rs.EOF
		If rs("badge_count") <> 0 Then
			iColorCount = iColorCount + 1
			
			If (iColorCount mod 2) Then
				sColorCode = "white"
			Else
				sColorCode = "#DFEFFF"
			End If
			
			sTableString2 = sTableString2 & "<tr bgcolor='" & sColorCode & "'><td width='300' align='left'>" & rs("country") & "</td>" & _
										    "<td width='100' align=""right"">" & rs("badge_count") & "</td></tr>"
											
			iCount2 = iCount2 + rs("badge_count")
		End If
		
		rs.MoveNext	
	Loop
	
	sTableString2 = sTableString2 & "<tr bgcolor='#AAD5FF'><td width='300' align='left'><b>Total</b></td>" & _
								    "<td width='100' align=""right"">" & iCount2 & _
								    "</td></tr><tr><td colspan='2' align='center'><a href='report_select.asp'>Return to Select Reports</a></td></tr></table>"
	
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
		<%=sTableString2%>
	</body>
</html>
