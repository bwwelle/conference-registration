<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim sBadgeString
	dim lngRecs
	dim strSQL
	dim conDB
	dim rs
	dim sPayType
	dim sAmount
	dim sProgramAmount
	dim sName
	dim sCompany
	dim sCompany1
	dim sCompany2
	dim s
	dim t
	dim i
	dim intPrintBadge
	dim intPrintBadgeReceipt
	dim intPrintProgramReceipt
	dim intProgramAmount
	dim intTotalAmount
	dim bolPrintHeader
	dim badge_string_num
	dim programtest
	dim programtester

	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
		
	sBadgeString = sBadgeString & "<tr><td width='110' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>"
	sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
	sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
	sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
	sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>"
	sBadgeString = sBadgeString & "<tr><td colspan='3' align='left' style='font-size:16px'><b>REGISTRATION PAYMENT DETAILS:</b></td></tr>" 
	
	sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>ITEM</b></td>" 
	
	sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>DATE</b></td>"
	
	sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>"
	
	sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>Program(s)</b></td>"
	
	sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>" & Date & "</b></td>" 
	
	sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>$" & Session("Addl_Prog_Cost")& "</b></td></tr>"
	
	sBadgeString = sBadgeString & "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" 
	
	sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:16px'><b>$" & Session("Addl_Prog_Cost") & "</b></td></tr>" 

	sBadgeString = sBadgeString & "</table>"
	
%>
<html>
	<head>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body onLoad="javascript:window.print();javascript:window.close();">
		<form name="printform">
			<% =sBadgeString %>
		</form>
	</body>
</html>
