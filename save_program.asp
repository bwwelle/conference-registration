<% @Language=VBScript %>
<% Response.Buffer="true"%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim conDB
	dim strSQL
	dim lngRecs
	dim iserror
	dim sPAYMENT
	dim cPAYMENT_TYPE
	dim sADDL_PROG_COUNT
	dim sADDL_PROG_COST
	
	sPAYMENT      		= Trim(Request("PAYMENT"))
	cPAYMENT_TYPE		= Trim(Request("PAYMENT_TYPE"))
	sADDL_PROG_COUNT	= Trim(Request("ADDL_PROG_COUNT"))
	sADDL_PROG_COST		= Trim(Request("ADDL_PROG_COST"))	
	
	Session("Payment")     		= Trim(Request("PAYMENT"))
	Session("Payment_Type")		= Trim(Request("PAYMENT_TYPE"))
	Session("Addl_Prog_Count")	= Trim(Request("ADDL_PROG_COUNT"))
	Session("Addl_Prog_Cost")	= Trim(Request("ADDL_PROG_COST"))
	iserror = 0

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
		
	conDB.BeginTrans
	
	If sADDL_PROG_COUNT <> "" Then 
		strSQL = "INSERT INTO program (registration_id, num_programs, payment_type, acttimestamp) VALUES (" 
		
		strSQL = strSQL & "NULL, " & CInt(sADDL_PROG_COUNT) & ", "
		
		If cPAYMENT_TYPE = "" Then 
		  strSQL = strSQL & "NULL, GETDATE())"
		Else       
		  strSQL = strSQL & "'" & cPAYMENT_TYPE & "', GETDATE())"
		End If
	End If
	
	conDB.Execute strSQL, lngRecs, adCmdtext
	
	If lngRecs <> 1 Then
		iserror = iserror + 1
	End if		
		
	If iserror = 0 Then
		conDB.CommitTrans
	End If

	conDB.Close
	
	Set conDB = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration - Save Exhibitor</title>
		<script LANGUAGE="JavaScript">
			function printBadges()
			{
				jWindow = window.open('print_program_receipt.asp', 'page_a', 'height=800, width=400, left=0, top=0', false);
					
				jWindow.focus();			
				}
		</script>
	</head>
	<body>
		<table width="100%" border="0">
			<form name="initiate_print" method=post action="default.asp" onSubmit="return printBadges()">
								<tr>
					<script language="JavaScript">
					{
					  printBadges();
					  
					  document.initiate_print.submit();
					}
					</script>
				</tr>
			</form>
		</table>
	</body>
</html>
