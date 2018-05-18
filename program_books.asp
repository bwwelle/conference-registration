<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration - Program Books</title>
<SCRIPT language="JavaScript">
<!--
			function CalcProgCost()
			{
			  program.ADDL_PROG_COST.value = (program.ADDL_PROG_COUNT.value * 20);
			  
			  program.PAYMENT.value = program.ADDL_PROG_COST.value;
			  
			  program.TOTAL_AMOUNT_PAID.value = (program.ADDL_PROG_COST.value - program.PAYMENT.value);
			}
//-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
<!-- 
			function CalcAmountDue()
			{
				program.TOTAL_AMOUNT_PAID.value = (program.ADDL_PROG_COST.value) - program.PAYMENT.value);
			}
//-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
<!-- 
			function validateForm()
			{ 	
				var numTypesChecked = 0;
				
			// is one payment type selected
				for (var i = 0; i <= 2; i++)
				{
					if (program.PAYMENT_TYPE[i].checked)
					{
						numTypesChecked ++;
					}
				}
				
				if ((numTypesChecked < 1) && (program.PAYMENT.value) != "")
				{
					alert("WARNING: One Payment Type must be selected.");
					
					return false;
				}
				else if ((numTypesChecked > 1) && (program.PAYMENT.value) != "")
				{
					alert("WARNING: Only one Payment Type can be selected.");
					
					return false;
				}
				else if (isNaN(program.ADDL_PROG_COUNT.value))
				{
					alert("WARNING: Add'l Programs entry is not a number.");
					
					program.ADDL_PROG_COUNT.value = "";
					
					return false;
				}
				else if (isNaN(program.PAYMENT.value) && (program.PAYMENT.value) != "")
				{
					alert("WARNING: Payment entry is not a number.");
					
					program.PAYMENT.value = "";
					
					return false;
				}
				else
				{
					alert("Purchase Completed.");
					
					return true;
				} 
			}
//-->
</SCRIPT>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body onLoad="document.program.ADDL_PROG_COUNT.focus()">
		<form name="program" method=post onSubmit="return validateForm()" action="save_program.asp">
			<table border=0 width=100% cellspacing="0">
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<strong>
							Purchase Program Books
						</strong>
					</td>
					<td width=8%>
					</td>
					<td width=16%>
					</td>
					<td width=48%>
					</td>
				</tr>
				<tr>
					<td colspan="2" align="right" valign="middle">
						Add'l Programs&nbsp;
						<input type=text value="" name="ADDL_PROG_COUNT" size="2" onKeyUp="CalcProgCost()">
						&nbsp;&nbsp;&nbsp;&nbsp;Cost
					</td>
					<td width=8%>
						<input type=text value="" name="ADDL_PROG_COST" size=10 readonly="true">
					</td>
					<td width=16%>
					</td>
					<td width=48%>
					</td>
				</tr>
				<tr>
					<td width=12%>
					</td>
					<td width=16% valign="top">
						<div align="right">
							Payment
						</div>
					</td>
					<td width=8% valign="top">
						<input type=text value="<%=PAYMENT%>" name="PAYMENT" size=10 >
					</td>
					<td width=16% valign="top">
						<input type="checkbox" name="PAYMENT_TYPE" value="CC" >
						Credit Card <br>
						<input type="checkbox" name="PAYMENT_TYPE" value="CH" >
						Check <br>
						<input type="checkbox" name="PAYMENT_TYPE" value="CA" >
						Cash/Trav Check
					</td>
					<td width=48%>
					</td>
				</tr>
				<tr>
					<td width=12%>
					</td>
					<td width=16%>
						<div align="right">
							Total Amount Due
						</div>
					</td>
					<td width=8%>
						<input type=text value="<%=TOTAL_AMOUNT_PAID%>" name="TOTAL_AMOUNT_PAID" size=10 readonly="true">
					</td>
					<!--<td width=8%><input type=text value="0" name="TOTAL_AMOUNT_PAID" size=10 readonly="true"></td>-->
					<td width=16%>
					</td>
					<td width=48%>
					</td>
				</tr>
			</table>
			<table border=0 width="100%">
				<tr>
					<td colspan=3>
						<div align="center">
							<input type="submit" name="SUBMIT" value="Submit">
							<input type="reset" name="RESET" value="Reset">
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="3">
						<div align="center">
							<a href="default.asp">
								Return to Home Page
							</a>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
