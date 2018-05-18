<% @Language=VBScript %>
<% Response.Buffer="true"%>
<% Server.ScriptTimeout = 999%>

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
	dim rsBadgeInfo
	dim iTotalDue
	dim rsBadgePaymentInfo
	dim company_name 
	dim booth_count
	dim comp_badge_count
	dim badge_claimed_count
	dim compBadgeCounter
	dim amountOwed
	dim printline
	dim iPickupCounter
	dim Exibitor_ID
	Exibitor_ID = Session("EXHIBITOR_ID")
	
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
	'"Provider=SQLOLEDB; Data Source=; Initial Catalog=; User Id=; Password=;"
		
	'get the counts and compay names for all those exhibitors who have picked up at least one badge
	strSQL = "select count(distinct b.badge_id) badge_count, count(distinct bth.booth_ID) as booth_count, e.company_name company_name "
	strSQL = strSQL & "from badge b, log_badge_activity l, booth bth, exhibitor_guest eg, exhibitor e "
	strSQL = strSQL & "where b.badge_id = l.badge_id AND bth.exhibitor_id = e.exhibitor_ID "
	strSQL = strSQL & "and b.guest_id = eg.guest_id "
	strSQL = strSQL & "and e.exhibitor_id = eg.exhibitor_id "
	strSQL = strSQL & "and b.badge_type = 'E' "	
	strSQL = strSQL & "and b.badge_id in "
	strSQL = strSQL & "(select badge_id from log_badge_activity) and e.exhibitor_id =  " & Exhibitor_ID
	strSQL = strSQL & "group by company_name "
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	firstline = "<tr bgcolor=""white"">"	

	company_name= rs("company_name")

	booth_count = rs("booth_count")

	comp_badge_count = CInt(rs("booth_count")) * 6
	
	'badge_claimed_count = rs("badge_count")	
	
	Set rs = Nothing	
	
	strSQL = "select distinct isnull(convert(varchar(50),l.acttimestamp,101),'Not Picked Up') as pickupdate, Case When bs.badge_ID is null then 'Show' else 'Preregistered' end as preregistered, eg.last_name as last_name, eg.first_name as first_name, b.badge_num as badge_num, b.guest_id, b.badge_id " & _
				 "from exhibitor_guest eg  inner join exhibitor e on e.exhibitor_id = eg.exhibitor_id Left outer join badges bs on bs.exhibitor_id = e.exhibitor_id_old and bs.lastname=eg.last_name and bs.firstname= eg.first_name, badge b LEFT OUTER JOIN log_badge_activity l ON b.badge_id = l.badge_id " & _
				 "where eg.guest_id = b.guest_id " & _
				 "and eg.exhibitor_id = " & Exhibitor_ID & " and eg.status = 'A'"	
				 
	Set rsBadgeInfo = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sTableString = ""
	iTotalDue = "0"	

	printline = 0
	iAmountOwed = 0
	compBadgeCounter = 0
	badge_claimed_count = 0
	iPickupCounter = 1
	
	Do While Not rsBadgeInfo.EOF
		If compBadgeCounter = CInt(comp_badge_count) then
			If printline = 0 then
				printline = 1
				
				sTableString = sTableString & "<tr><td align=""center"" colspan = 6><hr></td></tr>"			
			End if 
		End If
		
		If compBadgeCounter < CInt(comp_badge_count) OR rsBadgeInfo("pickupdate") <> "Not Picked Up" then
			sTableString = sTableString & "<tr><td align=""left"">" & iPickupCounter & ".)</td>"
			
			iPickupCounter = iPickupCounter + 1
		Else
			sTableString = sTableString & "<tr><td align=""left"">&nbsp;</td>"
		End If
		
		sTableString = sTableString & "<td align=""center"">" & rsBadgeInfo("last_name") & ", " & rsBadgeInfo("first_name") & "</td>"
		
		sTableString = sTableString & "<td align=""center"">" & rsBadgeInfo("preregistered") & "</td>"
				
		sTableString = sTableString & "<td align=""center"">" & rsBadgeInfo("pickupdate") & "</td>"
		
		If compBadgeCounter < CInt(comp_badge_count) then
			sTableString = sTableString & "<td align=""center"">Complimentary</td>"
			
			badge_claimed_count = badge_claimed_count + 1
		ElseIf rsBadgeInfo("pickupdate") = "Not Picked Up" then
			sTableString = sTableString & "<td align=""center"">$0</td>"
		ElseIf rsBadgeInfo("preregistered") = "Show" then	
			sTableString = sTableString & "<td align=""center"">$100</td>"
			
			iAmountOwed = iAmountOwed + 100
			
			badge_claimed_count = badge_claimed_count + 1
		ElseIf rsBadgeInfo("preregistered") = "Preregistered" then	
			sTableString = sTableString & "<td align=""center"">$90</td>"
			
			iAmountOwed = iAmountOwed + 90
			
			badge_claimed_count = badge_claimed_count + 1
		End If
		
		compBadgeCounter = compBadgeCounter + 1
		
		rsBadgeInfo.Movenext
	Loop

	set rsBadgeInfo = nothing
	
	dim iAmountPaid
	
	iAmountPaid = 0
	
	strSQL = "Select amount_paid from log_badge_payment_activity where exhibitor_id = 18175" 
				 
	Set rsBadgePaymentInfo = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
	Do While Not rsBadgePaymentInfo.EOF
		iAmountPaid = iAmountPaid + rsBadgeInfo("amount_paid")	
	Loop	
	
	set rsBadgePaymentInfo = nothing
	
	conDB.Close
	Set conDB = Nothing%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>The Registration - Badge Report</title>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body>
        <table width="100%"  border="0">
          <tr>
            <td colspan="2"><div align="right">
            Company Name: 
            </div></td>
            <td colspan="4"><%=company_name%></td>
          </tr>
          <tr>
            <td colspan="2"><div align="right">
            # Of Booths: 
            </div></td>
            <td colspan="4"><%=booth_count%></td>
          </tr>
          <tr>
            <td colspan="2"><div align="right">
            # Of Complimentary Badges:
            </div></td>
            <td colspan="4"><%=comp_badge_count %></td>
          </tr>
          <tr>
            <td colspan="2"><div align="right">
            Total Badges Picked Up: 
            </div></td>
            <td colspan="4"><%=badge_claimed_count %></td>
          </tr>
          <tr>
            <td colspan="6">&nbsp;</td>
          </tr>
          <tr>
		  	<td><div align="center">
            &nbsp;
            </div></td>
            <td><div align="center"><b>
            Name</b>
            </div></td>
            <td><div align="center">
            <b>Badge Added</b>
            </div></td>
            <td><div align="center">
            <b>Date Badge Picked Up</b>
            </div></td>
            <td><div align="center">
            <b>Badge Price</b> 
            </div></td>
          </tr>
		  <%= sTableString %>
          <tr>
            <td colspan="6">&nbsp;</td>
          </tr>
		  <tr>
            <td colspan="6"><hr></td>
          </tr>
		  <tr>
            <td colspan="2"><div align="right">
            SubTotal Amount Due:
            </div></td>
            <td colspan="4">$<%= iAmountOwed %></td>
          </tr>
		  	<tr>
            <td colspan="2"><div align="right">
            Total Amount Paid:
            </div></td>
            <td colspan="4">$<%= iAmountPaid %></td>
          </tr>
          <tr>
            <td colspan="2"><div align="right">
            <b>Total Amount Due:</b>
            </div></td>
            <td colspan="4"><b>$<%= iAmountOwed - iAmountPaid %></b></td>
          </tr>
        </table>	
  </body>
</html>
