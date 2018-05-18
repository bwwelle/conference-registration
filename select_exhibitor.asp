<% @Language=VBScript %>
<% response.buffer=true %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration Exhibitor</title>
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
		<form name="select_exhibitor" method=post action="select_exhibitor.asp">
			<table border=0>
				<tr>
					<td width=300>
						<div align="right">
							Search Company By Name
						</div>
					</td>
					<td width=300>
						<div align="left">
							<input type="text" name="SEARCH_COMPANY_NAME" size="50" value="">
						</div>
					</td>
					<td width=100>
						<div align="left">
							<input type="submit" name="SEARCH" value="Search">
						</div>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							--or--
						</div>
					</td>
					<td width="100">&nbsp;
					</td>
				</tr>
				<tr>
					<td width=300>
						<div align="right">
							Search Company By Guest Last Name
						</div>
					</td>
					<td width=300>
						<div align="left">
							<input type="text" name="SEARCH_NAME" size="50" value="">
						</div>
					</td>
					<td width=100>
						<div align="left">
							<input type="submit" name="SEARCHNAME" value="Search">
						</div>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<%
						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' search by company name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							dim conDB
							dim strSQL
							dim rsNames
							dim indexName
							dim lngRecs
							dim sSearchName
sSearchName = Replace(Trim(Request("SEARCH_NAME")),"'","''")

							Set conDB = Server.CreateObject("ADODB.Connection")
							
							conDB.Open strCon
							
							Session("COMPANY_LIST") = ""
							
							Session("EXHIBITOR_ID") = ""
							
							Session("SORT_ORDER") = "last_name,first_name"
							
							If Len(Request.form("SEARCH")) Then
								strSQL = "SELECT company_name FROM exhibitor " & _
										 "WHERE company_name like '" & Trim(Request("SEARCH_COMPANY_NAME")) & "%' " & _
										 "and status = 'A' " & _
										 "ORDER BY company_name"
										 
								Set rsNames = conDB.Execute(strSQL, lngRecs, adCmdtext)
								
								Response.Write("<select name=""COMPANY_LIST"" size=""20"" ONCHANGE=""document.select_exhibitor.submit(); return false;"">")
								
								Do While Not rsNames.EOF  
									Response.Write ("<option>" & rsNames("company_name") & "</option>")
									rsNames.MoveNext
								Loop
								
								Set rsNames = Nothing
								
								Response.Write("</div></select>")
							ElseIf Len(Request.form("SEARCHNAME")) Then
								strSQL = "SELECT eg.last_name last_name, eg.first_name first_name, e.company_name company_name " & _
										 "FROM exhibitor e, exhibitor_guest eg " & _
										 "WHERE e.exhibitor_id = eg.exhibitor_id " & _
										 "and e.status = 'A' " & _
										 "and eg.status = 'A' " & _
										 "and eg.last_name like '" & sSearchName & "%' " & _
										 "ORDER BY last_name, first_name "
										 
								Set rsNames = conDB.Execute(strSQL, lngRecs, adCmdtext)
								
								Response.Write("<select name=""GUEST_LIST"" size=""20"" ONCHANGE=""document.select_exhibitor.submit(); return false;"">")
								
								If rsNames.EOF Then
									Response.Write ("<option>No Guest Found</option>")
								Else
									Do While Not rsNames.EOF  
										Response.Write ("<option>" & rsNames("last_name") & ", " & rsNames("first_name") & ", " & rsNames("company_name") & "</option>")
										
										rsNames.MoveNext
									Loop
								End If
								
								Set rsNames = Nothing
								
								Response.Write("</div></select>")
							ElseIf Len(Request.form("COMPANY_LIST")) Then 
								Session("COMPANY_LIST") = Request.form("COMPANY_LIST")
								
								Response.Redirect("exhibitor.asp")
							ElseIf Len(Request.form("GUEST_LIST")) Then 
'first parse out the company name								
								arGuestField = split(Request.form("GUEST_LIST"), ", ")
								
								sLast = arGuestField(0)
								
								sFirst = arGuestField(1)
								
								sCompany = right(Request.form("GUEST_LIST"), len(Request.form("GUEST_LIST")) - (len(sLast) + len(sFirst) + 4))		
								
								Session("COMPANY_LIST") = sCompany
								
								Response.Redirect("exhibitor.asp")								
							ElseIf Len(Request.form("NEW")) Then 
								Response.Redirect("exhibitor_new.asp")
							End if				
							
							conDB.Close
							
							Set conDB = Nothing
						%>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							--or--
						</div>
					</td>
					<td width="100">&nbsp;
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							<input type="submit" name="NEW" value="Add New Exhibitor">
						</div>
					</td>
					<td width="100">&nbsp;
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
						<a href="default.asp">
							Return to Registration
						</a>
					</td>
					<td width="100">&nbsp;
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
