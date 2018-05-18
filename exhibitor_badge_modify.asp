<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsBadgeNames
	dim lngAllowedComp

'''''determine and set the comp badges based on booth number
	lngAllowedComp = (Session("NUM_BOOTHS") * 6)
	
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon	
	
'use exhibitor_id to get badge names
	strSQL = "select eg.first_name, eg.last_name, b.guest_id, b.badge_num as badge_num "
	
	strSQL = strSQL & "from exhibitor_guest eg, badge b "
	
	strSQL = strSQL & "where eg.guest_id = b.guest_id "
	
	strSQL = strSQL & "and eg.exhibitor_id = " & Session("EXHIBITOR_ID") & " and eg.status = 'A'"
	
	strSQL = strSQL & " order by b.badge_num"
	'response.write strSQL
	Set rsBadgeNames = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	Response.write("<form name=""exhibitor_badge_modify"" method=post action=""save_badge.asp"" onSubmit=""return validateForm()"">")
	
	Response.Write("<table width=""100%"" border=0 cellspacing=""0""><tr><td colspan=3><div align=""left""><a href=""exhibitor.asp"">Return to Exhibitor Information</a></div></td>")
	
	Response.Write("<td colspan=4><div align=""left""><input type=""submit"" name=""btnSubmit"" value=""Save Badges""></div></td></tr>")

	Response.Write("<tr bgcolor=""#AAD5FF"">" & _
	               "<td width=""7%""><b>Badge#</b></td>" & _
	               "<td colspan=2 align=""center""><b>First Name</b></td>" & _
				   "<td colspan=2 align=""center""><b>Last Name</b></td>" & _
				   "<td width=""17%"">Delete</td>" & _
	               "<td width=""17%""><input type=""hidden"" name=""NUM_COMP_BADGES_ALLOWED"" size=""20"" value=""" & (Session("NUM_BOOTHS") * 6) & """ readonly=""true""></td>" & _
				   "</tr>")
				   
'populate the form if there are badges, if not, display the empty rows   
	For x = 1 to 130
		If Not rsBadgeNames.EOF Then
			Response.Write("<tr>" & _
			               "<td width=""7%""><b>" & x & "</b></td>" & _
			               "<td width=""10%"" align=""right"">First Name</td>" & _
						   "<td width=""17%""><input type=""text"" name=""FIRSTNAME" & x & """ size=""20"" value=""" & rsBadgeNames("first_name") & """></td>" & _
						   "<td width=""10%"" align=""right"">Last Name</td>" & _
						   "<td width=""17%""><input type=""text"" name=""LASTNAME" & x & """ size=""20"" value=""" & rsBadgeNames("last_name") & """></td>" & _
						   "<td width=""17%""><input type=""checkbox"" name=""chkBadge"" value=""" & rsBadgeNames("guest_id") & """ ></td>" & _
						   "<td width=""17%""><input type=""hidden"" name=""GUESTID" & x & """ size=""20"" value=""" & rsBadgeNames("guest_id") & """ readonly=""true""></td>" & _
						   "</tr>") 
						   
			rsBadgeNames.Movenext
		Else
			Response.Write("<tr>" & _
			               "<td width=""7%""><b>" & x & "</b></td>" & _
			               "<td width=""10%"" align=""right"">First Name</td>" & _
						   "<td width=""17%""><input type=""text"" name=""FIRSTNAME" & x & """  value="""" size=""20""></td>" & _
						   "<td width=""10%"" align=""right"">Last Name</td>" & _
						   "<td width=""17%""><input type=""text"" name=""LASTNAME" & x & """  value="""" size=""20""></td>" & _
						   "<td width=""17%""><input type=""checkbox"" name=""chkBadge"" value="""" ></td>" & _						   
						   "<td width=""17%""><input type=""hidden"" name=""GUESTID" & x & """  value="""" size=""20"" readonly=""true""></td>" & _
						   "</tr>") 
		End If
	Next 
	
	Set rsBadgeNames = Nothing
	
	conDB.Close
	
	Set conDB = Nothing
	
	Response.Write("<tr><td colspan=3><div align=""left""><a href=""exhibitor.asp"">Return to Exhibitor Information</a></div></td>")
	
	Response.Write("<td colspan=4><div align=""left""><input type=""submit"" name=""btnSubmit"" value=""Save Badges""></div></td></tr>")
	
	Response.Write("</table></form>")
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
		<SCRIPT LANGUAGE="JavaScript">
			function validateForm()
			{
//total num of basdge for exhibitor is (number of booths times 6) + 6 (for extras)
				var numTotalBadges = parseInt(exhibitor_badge_modify.NUM_COMP_BADGES_ALLOWED.value) + 20;
				
				var numCount = 0;
				
				for (var i = 1; i <= numTotalBadges; i++)
				{
					if (eval("exhibitor_badge_modify.FIRSTNAME" + i + ".value.length") != 0)
					{
//the actual number of badges
						numCount ++;
					}
				}
				
				for (var i = 0; i <= numCount - 1; i++)
				{
					if (exhibitor_badge_modify.chkBadge[i].checked)
					{
						if (confirm("Delete badges?"))
						{
							return true;
						}
						else
						{
							return false;
						}
					}
				}
				
				if (confirm("There are " + numCount + " total badges. " + exhibitor_badge_modify.NUM_COMP_BADGES_ALLOWED.value + " are complimentary for this Exhibitor. Save badges?"))
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		</SCRIPT>		
	</head>	
	<body>	
	</body>
</html>
