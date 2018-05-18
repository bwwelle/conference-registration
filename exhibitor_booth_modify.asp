<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsExhibitor
	dim sCompany_name
	dim sBoothNums
	dim sBoothLocs
	
	strSQL = "select company_name from exhibitor where exhibitor_id = " & Session("EXHIBITOR_ID")
	
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
	
	Set rsExhibitor = conDB.Execute(strSQL, lngRecs, adCmdtext)	
	
	sCompany_name = rsExhibitor("company_name")
	
	Set rsExhibitor = Nothing
	
'''''booth information
'''''booth location info
	strSQL = "select booth_number from booth " & _
			 "where exhibitor_id = " & Session("EXHIBITOR_ID") & _
			 " and status = 'A' " & _
			 "order by booth_number"
			 
	Set rsBooths = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
'sBoothNums = ""
'sBoothLocs = ""
''''''''''''''''''''''''''''''''''''
	Do While Not rsBooths.EOF
			sBoothTable = sBoothTable & "<tr><td width=""23%"" align=""right""><input type=""checkbox"" name=""chkBooth"" value=""" & rsBooths("booth_number") & """ ></td>" & _
									"<td width=""2%"">&nbsp;</td>" & _
									"<td width=""10%"" align=""left"">" & rsBooths("booth_number") & "</td>" & _
									"<td width=""2%"">&nbsp;</td></tr>"

			rsBooths.Movenext
	Loop

	Set rsBooths = Nothing
	
	conDB.Close
	
	Set conDB = Nothing	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>The Registration Booth Information</title>
<style type="text/css">
		body,td,th 
		{
			font-family: Arial, Helvetica, sans-serif;
			font-size: x-small;
		}
		</style>
<SCRIPT LANGUAGE="JavaScript">
			function validateBooths()
			{
				var boothNum = parseInt(exhibitor_booth_modify.BOOTH_NUMBER.value);
				
				if (isNaN(boothNum) && (exhibitor_booth_modify.BOOTH_NUMBER.value) != "")
				{
					alert("Entry is not a number.");
					
					exhibitor_booth_modify.BOOTH_NUMBER.value = "";
					
					return false;
				}					
			}
		</SCRIPT>
</head>
<body>
<!-- BEGIN CONTENT -->
<form name="exhibitor_booth_modify" method=post onSubmit="return validateBooths()" action="save_booth.asp">
    <table border=0 width=100% cellspacing="0">
        <tr bgcolor="#AAD5FF">
            <td colspan=4>
                <strong>Booth Information</strong>
            </td>
        </tr>
        <tr>
            <td colspan=4 align="left">
                <strong>Company:&nbsp;<%=sCompany_name%></strong>
            </td>
        </tr>
        <tr>
            <td width="23%">
                <a href="exhibitor.asp">Back to Exhibitor Page</a>
            </td>
            <td colspan=3>&nbsp;
            </td>
        </tr>
        <tr>
            <td colspan = 4>&nbsp;
            </td>
        </tr>
        <tr>
            <td width="23%" align="right">
                <b>Delete</b>
            </td>
            <td width="2%" >&nbsp;
            </td>
            <td width="10%" >
                <b>Booth Number</b>
            </td>
            <td >&nbsp;
                            </td>
        </tr>
        <%=sBoothTable%>
        <tr>
            <td colspan = 4>&nbsp;
            </td>
        </tr>
        <tr>
            <td width="23%" align="right">
                <strong>Add Booth</strong>
            </td>
            <td width="2%" >&nbsp;
            </td>
            <td colspan=2 align="left">
                <input type=text value="" name="BOOTH_NUMBER" size=8>
                <font color="#FF0000">&nbsp;&nbsp;&nbsp;<%=Session("BOOTH_IS_USED")%></font>
            </td>
        </tr>
        <tr>
        	<td width="23%">
                <a href="exhibitor.asp">Back to Exhibitor Page</a>
            </td>
            <td width="2%" >&nbsp;
            </td>
            <td colspan=2 align="left">
                <input type="submit" name="btnSubmit" value="Submit">
            </td>
        </tr>
    </table>
</form>
</body>
</html>
