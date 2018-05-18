<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%

%>
<html>
	<head>
		<style type="text/css">
<!-- 
body,td,th {
	font-family: Arial Narrow;
}
-->
</style>
		<script LANGUAGE="JavaScript">
<!-- 
function DisplayAmountDue(btype)
{
	//alert("btype is " + btype);
	if (btype == "Director/Adult")
	{
		printform.PAID.value = "75.00";
	}
	else if (btype == "Commercial")
	{
		printform.PAID.value = "75.00";
	}
	else if (btype == "Guest")
	{
		printform.PAID.value = "75.00";
	}
	else if (btype == "College Student")
	{
		printform.PAID.value = "25.00";
	}
	else if (btype == "Clinician")
	{
		printform.PAID.value = "0.00";
	}
	else if (btype == "Staff")
	{
		printform.PAID.value = "0.00";
	}
}
//-->
</script>
	</head>
	<body >
		<form name="printform" method="post" action="manual_print_exhibitor_badge.asp">
			<table width="100%" border=0 cellspacing="0">
				<tr>
					<td colspan="2" align="center">
						<b>
							Manual Exhibitor Badge Print
						</b>
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">&nbsp;
					</td>
				</tr>
				<tr>
					<td width="50%" align="right">
						First Name/Last Name
					</td>
					<td width="50%" align="left">
						<input type="text" name="NAME" size="50">
					</td>
				</tr>
				<tr>
					<td width="50%" align="right">
						Company
					</td>
					<td width="50%" align="left">
						<input type="text" name="COMPANY" size="50">
					</td>
				</tr>
				<tr>
					<td width="50%" align="right">
						Amount Paid
					</td>
					<td width="50%" align="left">
						<input type="text" value="<%=PAID%>" name="PAID" size="10">
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">&nbsp;
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
						<input type="submit" name="SUBMIT" value="PRINT">
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
						<a href="default.asp">
							Return to Select Manual Badge Print
						</a>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
