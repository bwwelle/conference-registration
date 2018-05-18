<% @Language=VBScript %>
<% Response.Buffer="true"%>
<% Server.ScriptTimeout = 999%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sATTENDEE_LIST
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsAttendee
	dim rsProgID
	dim rsInitials
	dim sFirst
	dim sLast
	dim sCity
	dim sState
	dim sRegID
	dim sBadgeID
	dim arAttendeeField
	dim sPre
	dim sPreregistered
	dim sInitials
	dim sDisplayInitials
	dim iCnt
	
	Session("PRINT_ATTENDEE") = ""
	Session("DUPLICATE") = ""
	
	If Len(Request.form("PREVIOUS")) Or Len(Request.form("DUPLICATE")) Then
'we want to display the previous attendee - make it the current attendee
		Session("ATTENDEE_LIST") = Session("PREVIOUS_ATTENDEE")
	End If
	
	If Len(Request.form("DUPLICATE")) Then
'this is where the behavior of loading of the previous info changes -
'the form should load blank if a duplicate was previously saved
		Session("DUPLICATE") = "DUPLICATE"
	End If

	If Session("ATTENDEE_LIST") <> "" Then
'parse for first and last name, city and state
		arAttendeeField = split(Session("ATTENDEE_LIST"), " | ")
		
		sLast = arAttendeeField(0)
		sFirst = arAttendeeField(1)
		
		If arAttendeeField(2) = "|" then
			sCity = ""
			
			sState = ""
		ElseIf arAttendeeField(2) = "" or arAttendeeField(2) = null Then
			sCity = ""
			
			If arAttendeeField(3) = "" or arAttendeeField(3) = null Then
				sState = ""
			Else
				sState = arAttendeeField(3)
			End If
		ElseIf Right(arAttendeeField(2),1) = "|" Then		
			sCity = Mid(arAttendeeField(2),1,Len(arAttendeeField(2)) - 1)
			
			sState = ""
		Else
			sCity = arAttendeeField(2)
			
			If arAttendeeField(3) = "" or arAttendeeField(3) = null Then
				sState = ""
			Else
				sState = arAttendeeField(3)
			End If
		End If
		
		sPre = ""
		
		If InStr(1, sLast, "**") Then
			sPre = "*"
			
			sLast = Trim(Mid(sLast, 3))
		End If
		
		If InStr(1, sState, "^") Then
			sState = Trim(Left(sState, Len(sState)-1))
		End If
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
'get the requested attendee info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
		Set conDB = Server.CreateObject("ADODB.Connection")
		
		conDB.Open strCon
		
		If sPre = "*" Then
			strSQL = "SELECT b.badge_id, r.registration_id, isnull(r.share_info, '') share_info, isnull(r.zip, '') zip,  "
			
			strSQL = strSQL & "isnull(r.country, '') country, isnull(r.address, '') address, isnull(r.email, '') email, "
			
			strSQL = strSQL & "isnull(b.badge_city, '') city, isnull(b.badge_state, '') state, isnull(b.badge_type, '') badge_type, "
			
			strSQL = strSQL & "isnull(b.complimentary, 'N') complimentary, isnull(r.grade_level_type, '') grade_level_type, "
			
			strSQL = strSQL & "isnull(r.position_type, '') position_type, isnull(r.amount_paid, 0) amount_paid, isnull(r.amount_due, 0) amount_due, "
			
			strSQL = strSQL & "isnull(r.preregistered, '') preregistered, isnull(r.hotel, '') hotel, isnull(r.school_company, '') school_company, isnull(r.cell_phone, '') cell_phone, isnull(r.emergency_phone, '') emergency_phone, isnull(r.comments,'') comments "
			
			strSQL = strSQL & "FROM registration r, badge b "
			
			strSQL = strSQL & "WHERE r.registration_id = b.registration_id "
			
			strSQL = strSQL & "AND badge_first_name = '" & ManageQuotes(Trim(sFirst)) & "' "
			strSQL = strSQL & "AND badge_last_name = '" & ManageQuotes(Trim(sLast)) & "' "
			
			If sCity <> "" and Len(sCity)then 
				strSQL = strSQL & "AND badge_city = '" & ManageQuotes(Trim(sCity)) & "' "
			End If
			
			If sState <> "" and sState <> null then
				strSQL = strSQL & "AND badge_state = '" & ManageQuotes(Trim(sState)) & "' "	
			End If 	
		Else	
			strSQL = "SELECT b.badge_id, r.registration_id, isnull(r.share_info, '') share_info, isnull(r.zip, '') zip,  "
			
			strSQL = strSQL & "isnull(r.country, '') country, isnull(r.address, '') address, isnull(r.email, '') email, "
			
			strSQL = strSQL & "isnull(r.city, '') city, isnull(r.state, '') state, isnull(b.badge_type, '') badge_type, "
			
			strSQL = strSQL & "isnull(b.complimentary, 'N') complimentary, isnull(r.grade_level_type, '') grade_level_type, "
			
			strSQL = strSQL & "isnull(r.position_type, '') position_type, isnull(r.amount_paid, 0) amount_paid, isnull(r.amount_due, 0) amount_due, "
			
			strSQL = strSQL & "isnull(r.preregistered, '') preregistered, isnull(r.hotel, '') hotel, isnull(r.school_company, '') school_company, isnull(r.cell_phone,'') cell_phone,isnull(r.emergency_phone,'') emergency_phone, isnull(r.comments,'') comments "
			
			strSQL = strSQL & "FROM registration r, badge b "
			
			strSQL = strSQL & "WHERE r.registration_id = b.registration_id "
			
			strSQL = strSQL & "AND first_name = '" & ManageQuotes(Trim(sFirst)) & "' "
			
			strSQL = strSQL & "AND last_name = '" & ManageQuotes(Trim(sLast)) & "' "
			
			If sCity <> "" and Len(sCity) then 
				strSQL = strSQL & "AND city = '" & ManageQuotes(Trim(sCity)) & "' "
			End If
			
			If sState <> "" and sState <> null then
				strSQL = strSQL & "AND state = '" & ManageQuotes(Trim(sState)) & "' "	
			End If 
		End If

		Set rsAttendee = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		If lngRecs = -1 Then	
			If Session("DUPLICATE") <> "" Then
'set the session var so we know to insert instead of update
				FIRST_NAME			= ""
				EMAIL				= ""
				BADGE_TYPE			= ""
				BADGE_COMP			= ""
				GRADE_LEVELS		= ""
				AMOUNT_DUE			= "0"  
				PAYMENT				= "0"
				TO_BE_PAID			= "0"
				POSITION			= ""
				BADGE_INITIALS		= ""
				Session("BADGE_INITIALS") = ""
				SCHOOL_COMPANY  	= ""
				CELL_PHONE			= ""
				EMERGENCY_PHONE		= ""
				COMMENTS			= ""
			Else 
				FIRST_NAME			= sFirst
				EMAIL				= rsAttendee("email")
				BADGE_TYPE			= rsAttendee("badge_type")
				BADGE_COMP			= rsAttendee("complimentary")
				GRADE_LEVELS		= rsAttendee("grade_level_type")
				AMOUNT_DUE			= rsAttendee("amount_due")  'is the total reg cost
				PAYMENT				= rsAttendee("amount_paid")
				POSITION			= rsAttendee("position_type")
				BADGE_INITIALS		= Session("BADGE_INITIALS")
				SCHOOL_COMPANY  	= rsAttendee("school_company")
				CELL_PHONE			= rsAttendee("cell_phone")
				EMERGENCY_PHONE		= rsAttendee("emergency_phone")
				COMMENTS			= rsAttendee("comments")
			End If
	
			LAST_NAME		= sLast		
			SHARE_INFO		= rsAttendee("share_info")
			sBadgeID		= rsAttendee("badge_id")
			sRegID			= rsAttendee("registration_id")
			ZIP				= rsAttendee("zip")
			COUNTRY			= rsAttendee("country")
			ADDRESS			= rsAttendee("address")
			CITY			= rsAttendee("city")
			STATES			= rsAttendee("state")
			sPreregistered	= rsAttendee("preregistered")
			HOTEL 			= rsAttendee("hotel")			
		Else
'log it somewhere
		End if
	
		Set rsAttendee = Nothing
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
'check the extra programs, if any, get them
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		strSQL = "SELECT count(*) Cnt FROM program WHERE registration_id = " & sRegID 
		
		Set rsProgID = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		iCnt = rsProgID("Cnt")
		
		Set rsProgID = Nothing
		
		If iCnt <> 0 Then
			strSQL = "SELECT num_programs FROM program WHERE registration_id = " & sRegID
			
			Set rsProgID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			If Session("DUPLICATE") <> "" Then
				iCnt = 0
			Else
				iCnt = rsProgID("num_programs")
			End If
			
			Set rsProgID = Nothing
		End If
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' the payment is only for the registration so if there are
' programs, the payment for these has to be added
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
		iTotalPayment = (iCnt * 20) + CInt(PAYMENT)
		
		TO_BE_PAID = (CInt(AMOUNT_DUE) + (iCnt * 20)) - iTotalPayment 
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' the initials are in log_badge_payment_activity so get 
' them from there after rsAttendee is closed
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
		strSQL = "select max(acttimestamp) acttimestamp, initials, payment_type "
		
		strSQL = strSQL & "from log_badge_payment_activity "
		
		strSQL = strSQL & "where registration_id = " & sRegID
		
		strSQL = strSQL & " group by initials, payment_type"
		
		Set rsInitials = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
		If lngRecs = -1 then
			sInitials = rsInitials("initials")
			
			If Session("DUPLICATE") <> "" Then
				sPaymentType = ""
			Else
				sPaymentType = rsInitials("payment_type")
			End If
			
			Set rsInitials = Nothing
		
			If BADGE_INITIALS = "" Then
				sDisplayInitials = sInitials
			Else
				sDisplayInitials = BADGE_INITIALS
			End If
		Else
			sDisplayInitials = ""
			
			sPaymentType = "NC"
		End if
		
		conDB.Close
		
		Set conDB = Nothing
	Else	
		If Session("LOOKUP_CITY") <> "" Then
			CITY = trim(Session("LOOKUP_CITY"))
			
			STATES = trim(Session("LOOKUP_STATE"))
			
			ZIP = trim(Session("LOOKUP_ZIP"))
		Else
'place zipcode not found here?
		End If
		
		If COUNTRY = "" Then
			COUNTRY = "United States"
		End If
	
		Session("LOOKUP_CITY") = ""
		
		Session("LOOKUP_STATE") = ""
		
		Session("LOOKUP_ZIP") = ""	
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" encoding="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration</title>
		<script LANGUAGE="JavaScript">
			function IsItFive()
			{
// when the onKeyUp event is fired for the zip code entry,
// look for 5 numbers and then parse the xmldoc
				if (zcode.LOOKUPTHISZIP.value.length == 5)
				{
					document.zcode.submit();
				}
			}
		
			function CalcProgCost()
			{
			  if (registration.ADDL_PROG_COUNT.value == "")
			  {
				registration.ADDL_PROG_COUNT.value = "0";
			  }
			  
			  registration.ADDL_PROG_COST.value = (parseInt(registration.ADDL_PROG_COUNT.value) * 20);
			  
			  CalcAmountDue();
			}
		
			function CalcAmountDue()
			{
			  if (registration.ADDL_PROG_COST.value == "")
			  {
				registration.ADDL_PROG_COST.value = "0";
			  }
			  
			  registration.TOTAL_AMOUNT_DUE.value = parseInt(registration.AMOUNT_DUE.value) + parseInt(registration.ADDL_PROG_COST.value);
			  
			  registration.PAYMENT.value = parseInt(registration.TOTAL_AMOUNT_DUE.value);
			  
			  registration.TO_BE_PAID.value = parseInt(registration.TOTAL_AMOUNT_DUE.value) - parseInt(registration.PAYMENT.value);
			}
		
			function CalcToBePaid()
			{
			  registration.TO_BE_PAID.value = parseInt(registration.TOTAL_AMOUNT_DUE.value) - parseInt(registration.PAYMENT.value);
			}
		
			function NoChargeClicked()
			{
				if (registration.PAYMENT_TYPE[3].checked)
				{
					registration.BADGE_COMP.checked = true;
					
					DisplayAmountDue("Y");
				}
				else
				{
					registration.BADGE_COMP.checked = false;
					
					registration.PAYMENT_TYPE[0].disabled = false;
					
					registration.PAYMENT_TYPE[1].disabled = false;
					
					registration.PAYMENT_TYPE[2].disabled = false;
				}
			}
					
			function DisplayAmountDue(btype)
			{
			  if (btype == "Y")
			  {
				if (registration.BADGE_COMP.checked)
				{
					registration.PAYMENT_TYPE[0].checked = false;
					
					registration.PAYMENT_TYPE[1].checked = false;
					
					registration.PAYMENT_TYPE[2].checked = false;
					
					registration.PAYMENT_TYPE[3].checked = true;
				
					registration.PAYMENT_TYPE[0].disabled = true;
					
					registration.PAYMENT_TYPE[1].disabled = true;
					
					registration.PAYMENT_TYPE[2].disabled = true;
				
					registration.AMOUNT_DUE.value = "0";
					
					CalcProgCost();
				}
				else if (registration.BADGE_COMP.checked == false)
				{
					registration.PAYMENT_TYPE[3].checked = false;
					
					registration.PAYMENT_TYPE[0].disabled = false;
					
					registration.PAYMENT_TYPE[1].disabled = false;
					
					registration.PAYMENT_TYPE[2].disabled = false;
				}
			  }
			  else if (btype == "D")
			  {
				if (registration.BADGE_TYPE[0].checked)
				{
					if (registration.PREREG.value == "Y")
					{
					  registration.AMOUNT_DUE.value = "110";
					  
					  CalcProgCost();
					}
					else 
					{
					  registration.AMOUNT_DUE.value = "135";
					  
					  CalcProgCost();
					}		
				}
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "S")
			  {
				if (registration.BADGE_TYPE[1].checked)
				{
					if(registration.PREREG.value == "Y")
					{
						registration.AMOUNT_DUE.value = "50";
					}
					else
					{
						registration.AMOUNT_DUE.value = "75";
					}
					
					CalcProgCost();
				}
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "L")
			  {
				if (registration.BADGE_TYPE[2].checked)
					if (registration.PREREG.value == "Y")
					{
					  registration.AMOUNT_DUE.value = "110";
					  
					  CalcProgCost();
					}
					else
					{
					  registration.AMOUNT_DUE.value = "135";
					  
					  CalcProgCost();
					}		
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "C")
			  {
				if (registration.BADGE_TYPE[3].checked)
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "G")
			  {
				if (registration.BADGE_TYPE[4].checked)
					if (registration.PREREG.value == "Y")
					{
					  registration.AMOUNT_DUE.value = "110";
					  
					  CalcProgCost();
					}
					else
					{
					  registration.AMOUNT_DUE.value = "135";
					  
					  CalcProgCost();
					}		
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "F")
			  {
				if (registration.BADGE_TYPE[5].checked)
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			  else if (btype == "BS")
			  {
				if (registration.BADGE_TYPE[6].checked)
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
				else
				{
				  registration.AMOUNT_DUE.value = "0";
				  
				  CalcProgCost();
				}
			  }
			}
		 
			function InStr(strSearch, charSearchFor)
			{
				for (i=0; i < strSearch.Length; i++)
				{
					if (Mid(strSearch, i, 1)==charSearchFor)
					{
						return i;
					}
				}
				
				return -1;
			}
			
			function validateForm()
			{ 
				if (InStr(registration.FIRST_NAME.value,",")!=-1 || InStr(registration.LAST_NAME.value,",")!=-1 || InStr(registration.CITY.value,",")!=-1)
				{
					alert("The first name, last name, and city fields cannot contain a comma.");
					
					return false;
				}
				
				var numTypesChecked = 0;
				
// is only one badge_type selected
				for (var i = 0; i <= 6; i++)
				{
					if (registration.BADGE_TYPE[i].checked)
					{
						numTypesChecked ++;
					}
				}
				
				if (numTypesChecked > 1)
				{
					alert("Only one Badge Type may be selected.");
					
					return false;
				}
				else if (numTypesChecked == 0)
				{
					alert("A Badge Type must be selected.");
					
					return false;
				}	
				
// is one payment type selected
				numTypesChecked = 0;
				
				for (var i = 0; i <= 3; i++)
				{
					if (registration.PAYMENT_TYPE[i].checked)
					{
						numTypesChecked ++;
					}
				}
				
				if ((numTypesChecked < 1) && (registration.PAYMENT.value) != "0")
				{
					alert("One Payment Type must be selected.");
						
					return false;
				}
				else if ((numTypesChecked > 1) && (registration.PAYMENT.value) != "")
				{
					alert("Only one Payment Type can be selected.");
					
					return false;
				}
				// Check whether fields are empty.
				else if (registration.FIRST_NAME.value.length == 0)
				{
					alert("Please enter the First Name.");
					
					return false;
				}
				else if (registration.LAST_NAME.value.length == 0)
				{
					alert("Please enter the Last Name.");
					
					return false;
				}
				else if (registration.CITY.value.length == 0)
				{
					alert("Please enter the City.");
					
					return false;
				}
				else if (isNaN(registration.ADDL_PROG_COUNT.value) && (registration.ADDL_PROG_COUNT.value) != "")
				{
					alert("Add'l Programs entry is not a number.");
					
					registration.ADDL_PROG_COUNT.value = "";
					
					return false;
				}
				else if (isNaN(registration.PAYMENT.value) && (registration.PAYMENT.value) != "")
				{
					alert("Payment entry is not a number.");
					
					registration.PAYMENT.value = "";
					
					return false;
				}
				else
				{
					return true;
				} 
			}
			
			function setfocus() 
			{
				<%
					if (Session("SEARCH_FOR_ZIP") = "true")then
				Session("SEARCH_FOR_ZIP") = "false"
				%>
				
						document.zcode.LOOKUPTHISZIP.focus();
				<%
					else
				%>
						document.registration.FIRST_NAME.focus();
				<%
					end if
				%>	  
			}
		</script>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body onLoad="setfocus();">
		<!-- BEGIN CONTENT -->
		<table border=0 width="100%">
			<tr>
				<td width=50%>
					<form name="registration11" method=post action="registration.asp">
						<div align="right">
						</div>
					</form>
				</td>
				<td width=50%>
					<form name="registration1" method=post action="registration.asp">
						<div align="left">
						</div>
					</form>
				</td>
			</tr>
		</table>
		<table width="100%" border=0 cellspacing="0">
			<form name="zcode" method=post action="ziplookup.asp">
				<tr bgcolor="#AAD5FF">
					<td>
						<strong>
							ZipCode Look-Up
						</strong>
						<input type="text" value="<%=LOOKUPTHISZIP%>" name="LOOKUPTHISZIP" size="15" onKeyUp="IsItFive()" tabindex="1">
					</td>
				</tr>
			</form>
		</table>
		<form name="registration" id="registration" method=post onSubmit="return validateForm()" action="save_registration.asp">
			<table width="100%" border=0 cellspacing="0">
<!-- BEGIN REGISTRATION INFO -->
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<strong>
							Registration Information
						</strong>
					</td>
					<td width=85%>
						<strong>
							<font color="red">
								<% if sPreregistered="Y" then response.write (" *Preregistered* ")%>
							</font>
						</strong>
					</td>
					<td width=85%>
						<input type="hidden" name="PREREG" value="<%=sPreregistered%>">
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							First Name
						</div>
					</td>
					<td width=85%>
						<input type="text" value="<%=FIRST_NAME%>" id="FIRST_NAME" name="FIRST_NAME" size=50 maxlength=50 tabindex="2">
					</td>
					<td width=85%>
						<div align="right">
							Last Name
						</div>
					</td>
					<td width=85%>
                      <input type="text" value="<%=LAST_NAME%>" name="LAST_NAME" size=50 maxlength=50 tabindex="3">
                    </td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							School/Company
						</div>
					</td>
					<td colspan=3 width=85%>
						<input type="text" value="<%=SCHOOL_COMPANY%>" name="SCHOOL_COMPANY" size=125 maxlength=125 tabindex="5">
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Address
						</div>
					</td>
					<td colspan=3 width=85%>
						<input type=text value="<%=ADDRESS%>" name="ADDRESS" size=125 maxlength=125 tabindex="6">
					</td>
				</tr>
				<tr>
					<td width=15% height="26">
						<div align="right">
							City
						</div>
					</td>
					<td width=85%>
						<input value="<%=CITY%>" type=text name="CITY" size=50 maxlength=50 tabindex="7">
					</td>
					<td width=85%>
						<div align="right">
							State/Province/Region
						</div>
					</td>
					<td width=85%>
						<select name="STATES" tabindex="8">
							<option value="">
								&nbsp;
							</option>
							<option <% if STATES="AL" then response.write (" SELECTED ")%>>
								Alabama
							</option>
							<option <% if STATES="AK" then response.write (" SELECTED ")%>>
								Alaska
							</option>
							<option <% if STATES="AB" then response.write (" SELECTED ")%>>
								Alberta
							</option>
							<option <% if STATES="AZ" then response.write (" SELECTED ")%>>
								Arizona
							</option>
							<option <% if STATES="AR" then response.write (" SELECTED ")%>>
								Arkansas
							</option>
							<option <% if STATES="BC" then response.write (" SELECTED ")%>>
								British Columbia
							</option>
							<option <% if STATES="CA" then response.write (" SELECTED ")%>>
								California
							</option>
							<option <% if STATES="CO" then response.write (" SELECTED ")%>>
								Colorado
							</option>
							<option <% if STATES="CT" then response.write (" SELECTED ")%>>
								Connecticut
							</option>
							<option <% if STATES="DE" then response.write (" SELECTED ")%>>
								Delaware
							</option>
							<option <% if STATES="FL" then response.write (" SELECTED ")%>>
								Florida
							</option>
							<option <% if STATES="GA" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if STATES="HI" then response.write (" SELECTED ")%>>
								Hawaii
							</option>
							<option <% if STATES="ID" then response.write (" SELECTED ")%>>
								Idaho
							</option>
							<option <% if STATES="IL" then response.write (" SELECTED ")%>>
								Illinois
							</option>
							<option <% if STATES="IN" then response.write (" SELECTED ")%>>
								Indiana
							</option>
							<option <% if STATES="IA" then response.write (" SELECTED ")%>>
								Iowa
							</option>
							<option <% if STATES="KS" then response.write (" SELECTED ")%>>
								Kansas
							</option>
							<option <% if STATES="KY" then response.write (" SELECTED ")%>>
								Kentucky
							</option>
							<option <% if STATES="LA" then response.write (" SELECTED ")%>>
								Louisiana
							</option>
							<option <% if STATES="ME" then response.write (" SELECTED ")%>>
								Maine
							</option>
							<option <% if STATES="MB" then response.write (" SELECTED ")%>>
								Manitoba
							</option>
							<option <% if STATES="MD" then response.write (" SELECTED ")%>>
								Maryland
							</option>
							<option <% if STATES="MA" then response.write (" SELECTED ")%>>
								Massachusetts
							</option>
							<option <% if STATES="MI" then response.write (" SELECTED ")%>>
								Michigan
							</option>
							<option <% if STATES="MN" then response.write (" SELECTED ")%>>
								Minnesota
							</option>
							<option <% if STATES="MS" then response.write (" SELECTED ")%>>
								Mississippi
							</option>
							<option <% if STATES="MO" then response.write (" SELECTED ")%>>
								Missouri
							</option>
							<option <% if STATES="MT" then response.write (" SELECTED ")%>>
								Montana
							</option>
							<option <% if STATES="NE" then response.write (" SELECTED ")%>>
								Nebraska
							</option>
							<option <% if STATES="NV" then response.write (" SELECTED ")%>>
								Nevada
							</option>
							<option <% if STATES="NB" then response.write (" SELECTED ")%>>
								New Brunswick
							</option>
							<option <% if STATES="NH" then response.write (" SELECTED ")%>>
								New Hampshire
							</option>
							<option <% if STATES="NJ" then response.write (" SELECTED ")%>>
								New Jersey
							</option>
							<option <% if STATES="NM" then response.write (" SELECTED ")%>>
								New Mexico
							</option>
							<option <% if STATES="NSW" then response.write (" SELECTED ")%>>
								New South Wales
							</option>
							<option <% if STATES="NY" then response.write (" SELECTED ")%>>
								New York
							</option>
							<option <% if STATES="NL" then response.write (" SELECTED ")%>>
								Newfoundland
							</option>
							<option <% if STATES="NC" then response.write (" SELECTED ")%>>
								North Carolina
							</option>
							<option <% if STATES="ND" then response.write (" SELECTED ")%>>
								North Dakota
							</option>
							<option <% if STATES="NT" then response.write (" SELECTED ")%>>
								Northern Territory
							</option>
							<option <% if STATES="NT" then response.write (" SELECTED ")%>>
								Northwest Territories
							</option>
							<option <% if STATES="NS" then response.write (" SELECTED ")%>>
								Nova Scotia
							</option>
							<option <% if STATES="NU" then response.write (" SELECTED ")%>>
								Nunavut
							</option>
							<option <% if STATES="OH" then response.write (" SELECTED ")%>>
								Ohio
							</option>
							<option <% if STATES="OK" then response.write (" SELECTED ")%>>
								Oklahoma
							</option>
							<option <% if STATES="ON" then response.write (" SELECTED ")%>>
								Ontario
							</option>
							<option <% if STATES="OR" then response.write (" SELECTED ")%>>
								Oregon
							</option>
							<option <% if STATES="PA" then response.write (" SELECTED ")%>>
								Pennsylvania
							</option>
							<option <% if STATES="PE" then response.write (" SELECTED ")%>>
								Prince Edward Island
							</option>
							<option <% if STATES="QC" then response.write (" SELECTED ")%>>
								Quebec
							</option>
							<option <% if STATES="QLD" then response.write (" SELECTED ")%>>
								Queensland
							</option>
							<option <% if STATES="RI" then response.write (" SELECTED ")%>>
								Rhode Island
							</option>
							<option <% if STATES="SK" then response.write (" SELECTED ")%>>
								Saskatchewan
							</option>
							<option <% if STATES="SA" then response.write (" SELECTED ")%>>
								South Australia
							</option>
							<option <% if STATES="SC" then response.write (" SELECTED ")%>>
								South Carolina
							</option>
							<option <% if STATES="SD" then response.write (" SELECTED ")%>>
								South Dakota
							</option>
							<option <% if STATES="TAS" then response.write (" SELECTED ")%>>
								Tasmania
							</option>
							<option <% if STATES="TN" then response.write (" SELECTED ")%>>
								Tennessee
							</option>
							<option <% if STATES="TX" then response.write (" SELECTED ")%>>
								Texas
							</option>
							<option <% if STATES="UT" then response.write (" SELECTED ")%>>
								Utah
							</option>
							<option <% if STATES="VT" then response.write (" SELECTED ")%>>
								Vermont
							</option>
							<option <% if STATES="VIC" then response.write (" SELECTED ")%>>
								Victoria
							</option>
							<option <% if STATES="VA" then response.write (" SELECTED ")%>>
								Virginia
							</option>
							<option <% if STATES="WA" then response.write (" SELECTED ")%>>
								Washington
							</option>
							<option <% if STATES="DC" then response.write (" SELECTED ")%>>
								Washington DC
							</option>
							<option <% if STATES="WV" then response.write (" SELECTED ")%>>
								West Virginia
							</option>
							<option <% if STATES="W AUST" then response.write (" SELECTED ")%>>
								Western Australia
							</option>
							<option <% if STATES="WI" then response.write (" SELECTED ")%>>
								Wisconsin
							</option>
							<option <% if STATES="WY" then response.write (" SELECTED ")%>>
								Wyoming
							</option>
							<option <% if STATES="YT" then response.write (" SELECTED ")%>>
								Yukon Territory
							</option>
							<option <% if STATES="Alabama" then response.write (" SELECTED ")%>>
								Alabama
							</option>
							<option <% if STATES="Alaska" then response.write (" SELECTED ")%>>
								Alaska
							</option>
							<option <% if STATES="Alberta" then response.write (" SELECTED ")%>>
								Alberta
							</option>
							<option <% if STATES="Arizona" then response.write (" SELECTED ")%>>
								Arizona
							</option>
							<option <% if STATES="Arkansas" then response.write (" SELECTED ")%>>
								Arkansas
							</option>
							<option <% if STATES="British Columbia" then response.write (" SELECTED ")%>>
								British Columbia
							</option>
							<option <% if STATES="California" then response.write (" SELECTED ")%>>
								California
							</option>
							<option <% if STATES="Colorado" then response.write (" SELECTED ")%>>
								Colorado
							</option>
							<option <% if STATES="Connecticut" then response.write (" SELECTED ")%>>
								Connecticut
							</option>
							<option <% if STATES="Delaware" then response.write (" SELECTED ")%>>
								Delaware
							</option>
							<option <% if STATES="Florida" then response.write (" SELECTED ")%>>
								Florida
							</option>
							<option <% if STATES="Georgia" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if STATES="Hawaii" then response.write (" SELECTED ")%>>
								Hawaii
							</option>
							<option <% if STATES="Idaho" then response.write (" SELECTED ")%>>
								Idaho
							</option>
							<option <% if STATES="Illinois" then response.write (" SELECTED ")%>>
								Illinois
							</option>
							<option <% if STATES="Indiana" then response.write (" SELECTED ")%>>
								Indiana
							</option>
							<option <% if STATES="Iowa" then response.write (" SELECTED ")%>>
								Iowa
							</option>
							<option <% if STATES="Kansas" then response.write (" SELECTED ")%>>
								Kansas
							</option>
							<option <% if STATES="Kentucky" then response.write (" SELECTED ")%>>
								Kentucky
							</option>
							<option <% if STATES="Louisiana" then response.write (" SELECTED ")%>>
								Louisiana
							</option>
							<option <% if STATES="Maine" then response.write (" SELECTED ")%>>
								Maine
							</option>
							<option <% if STATES="Manitoba" then response.write (" SELECTED ")%>>
								Manitoba
							</option>
							<option <% if STATES="Maryland" then response.write (" SELECTED ")%>>
								Maryland
							</option>
							<option <% if STATES="Massachusetts" then response.write (" SELECTED ")%>>
								Massachusetts
							</option>
							<option <% if STATES="Michigan" then response.write (" SELECTED ")%>>
								Michigan
							</option>
							<option <% if STATES="Minnesota" then response.write (" SELECTED ")%>>
								Minnesota
							</option>
							<option <% if STATES="Mississippi" then response.write (" SELECTED ")%>>
								Mississippi
							</option>
							<option <% if STATES="Missouri" then response.write (" SELECTED ")%>>
								Missouri
							</option>
							<option <% if STATES="Montana" then response.write (" SELECTED ")%>>
								Montana
							</option>
							<option <% if STATES="Nebraska" then response.write (" SELECTED ")%>>
								Nebraska
							</option>
							<option <% if STATES="Nevada" then response.write (" SELECTED ")%>>
								Nevada
							</option>
							<option <% if STATES="New Brunswick" then response.write (" SELECTED ")%>>
								New Brunswick
							</option>
							<option <% if STATES="New Hampshire" then response.write (" SELECTED ")%>>
								New Hampshire
							</option>
							<option <% if STATES="New Jersey" then response.write (" SELECTED ")%>>
								New Jersey
							</option>
							<option <% if STATES="New Mexico" then response.write (" SELECTED ")%>>
								New Mexico
							</option>
							<option <% if STATES="New South Wales" then response.write (" SELECTED ")%>>
								New South Wales
							</option>
							<option <% if STATES="New York" then response.write (" SELECTED ")%>>
								New York
							</option>
							<option <% if STATES="Newfoundland" then response.write (" SELECTED ")%>>
								Newfoundland
							</option>
							<option <% if STATES="North Carolina" then response.write (" SELECTED ")%>>
								North Carolina
							</option>
							<option <% if STATES="North Dakota" then response.write (" SELECTED ")%>>
								North Dakota
							</option>
							<option <% if STATES="Northern Territory" then response.write (" SELECTED ")%>>
								Northern Territory
							</option>
							<option <% if STATES="Northwest Territories" then response.write (" SELECTED ")%>>
								Northwest Territories
							</option>
							<option <% if STATES="Nova Scotia" then response.write (" SELECTED ")%>>
								Nova Scotia
							</option>
							<option <% if STATES="Nunavut" then response.write (" SELECTED ")%>>
								Nunavut
							</option>
							<option <% if STATES="Ohio" then response.write (" SELECTED ")%>>
								Ohio
							</option>
							<option <% if STATES="Oklahoma" then response.write (" SELECTED ")%>>
								Oklahoma
							</option>
							<option <% if STATES="Ontario" then response.write (" SELECTED ")%>>
								Ontario
							</option>
							<option <% if STATES="Oregon" then response.write (" SELECTED ")%>>
								Oregon
							</option>
							<option <% if STATES="Pennsylvania" then response.write (" SELECTED ")%>>
								Pennsylvania
							</option>
							<option <% if STATES="Prince Edward Island" then response.write (" SELECTED ")%>>
								Prince Edward Island
							</option>
							<option <% if STATES="Quebec" then response.write (" SELECTED ")%>>
								Quebec
							</option>
							<option <% if STATES="Queensland" then response.write (" SELECTED ")%>>
								Queensland
							</option>
							<option <% if STATES="Rhode Island" then response.write (" SELECTED ")%>>
								Rhode Island
							</option>
							<option <% if STATES="Saskatchewan" then response.write (" SELECTED ")%>>
								Saskatchewan
							</option>
							<option <% if STATES="South Australia" then response.write (" SELECTED ")%>>
								South Australia
							</option>
							<option <% if STATES="South Carolina" then response.write (" SELECTED ")%>>
								South Carolina
							</option>
							<option <% if STATES="South Dakota" then response.write (" SELECTED ")%>>
								South Dakota
							</option>
							<option <% if STATES="Tasmania" then response.write (" SELECTED ")%>>
								Tasmania
							</option>
							<option <% if STATES="Tennessee" then response.write (" SELECTED ")%>>
								Tennessee
							</option>
							<option <% if STATES="Texas" then response.write (" SELECTED ")%>>
								Texas
							</option>
							<option <% if STATES="Utah" then response.write (" SELECTED ")%>>
								Utah
							</option>
							<option <% if STATES="Vermont" then response.write (" SELECTED ")%>>
								Vermont
							</option>
							<option <% if STATES="Victoria" then response.write (" SELECTED ")%>>
								Victoria
							</option>
							<option <% if STATES="Virginia" then response.write (" SELECTED ")%>>
								Virginia
							</option>
							<option <% if STATES="Washington" then response.write (" SELECTED ")%>>
								Washington
							</option>
							<option <% if STATES="Washington DC" then response.write (" SELECTED ")%>>
								Washington DC
							</option>
							<option <% if STATES="West Virginia" then response.write (" SELECTED ")%>>
								West Virginia
							</option>
							<option <% if STATES="Western Australia" then response.write (" SELECTED ")%>>
								Western Australia
							</option>
							<option <% if STATES="Wisconsin" then response.write (" SELECTED ")%>>
								Wisconsin
							</option>
							<option <% if STATES="Wyoming" then response.write (" SELECTED ")%>>
								Wyoming
							</option>
							<option <% if STATES="Yukon Territory" then response.write (" SELECTED ")%>>
								Yukon Territory
							</option>
						</select>
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Zip/Postal Code
						</div>
					</td>
					<td width=85%>
                      <input type="text" value="<%=ZIP%>" name="ZIP" size=15 tabindex="9">
                    </td>
				</tr>
				<tr>
					<td>
						<div align="right">
							Country
						</div>
					</td>
					<td>
						<select name="COUNTRY" tabindex="10">
							<option value="<%=COUNTRY%>">
								United States
							</option>
							<option <% if COUNTRY="Albania" then response.write (" SELECTED ")%>>
								Albania
							</option>
							<option <% if COUNTRY="Algeria" then response.write (" SELECTED ")%>>
								Algeria
							</option>
							<option <% if COUNTRY="American Samoa" then response.write (" SELECTED ")%>>
								American Samoa
							</option>
							<option <% if COUNTRY="Andorra" then response.write (" SELECTED ")%>>
								Andorra
							</option>
							<option <% if COUNTRY="Angola" then response.write (" SELECTED ")%>>
								Angola
							</option>
							<option <% if COUNTRY="Anguilla" then response.write (" SELECTED ")%>>
								Anguilla
							</option>
							<option <% if COUNTRY="Antarctica" then response.write (" SELECTED ")%>>
								Antarctica
							</option>
							<option <% if COUNTRY="Antigua And Barbuda" then response.write (" SELECTED ")%>>
								Antigua And Barbuda
							</option>
							<option <% if COUNTRY="Argentina" then response.write (" SELECTED ")%>>
								Argentina
							</option>
							<option <% if COUNTRY="Armenia" then response.write (" SELECTED ")%>>
								Armenia
							</option>
							<option <% if COUNTRY="Aruba" then response.write (" SELECTED ")%>>
								Aruba
							</option>
							<option <% if COUNTRY="Australia" then response.write (" SELECTED ")%>>
								Australia
							</option>
							<option <% if COUNTRY="Austria" then response.write (" SELECTED ")%>>
								Austria
							</option>
							<option <% if COUNTRY="Azerbaijan" then response.write (" SELECTED ")%>>
								Azerbaijan
							</option>
							<option <% if COUNTRY="Bahamas" then response.write (" SELECTED ")%>>
								Bahamas
							</option>
							<option <% if COUNTRY="Bahrain" then response.write (" SELECTED ")%>>
								Bahrain
							</option>
							<option <% if COUNTRY="Bangladesh" then response.write (" SELECTED ")%>>
								Bangladesh
							</option>
							<option <% if COUNTRY="Barbados" then response.write (" SELECTED ")%>>
								Barbados
							</option>
							<option <% if COUNTRY="Belarus" then response.write (" SELECTED ")%>>
								Belarus
							</option>
							<option <% if COUNTRY="Belgium" then response.write (" SELECTED ")%>>
								Belgium
							</option>
							<option <% if COUNTRY="Belize" then response.write (" SELECTED ")%>>
								Belize
							</option>
							<option <% if COUNTRY="Benin" then response.write (" SELECTED ")%>>
								Benin
							</option>
							<option <% if COUNTRY="Bermuda" then response.write (" SELECTED ")%>>
								Bermuda
							</option>
							<option <% if COUNTRY="Bhutan" then response.write (" SELECTED ")%>>
								Bhutan
							</option>
							<option <% if COUNTRY="Bolivia" then response.write (" SELECTED ")%>>
								Bolivia
							</option>
							<option <% if COUNTRY="Bosnia and Herzegovina" then response.write (" SELECTED ")%>>
								Bosnia and Herzegovina
							</option>
							<option <% if COUNTRY="Botswana" then response.write (" SELECTED ")%>>
								Botswana
							</option>
							<option <% if COUNTRY="Bouvet Island" then response.write (" SELECTED ")%>>
								Bouvet Island
							</option>
							<option <% if COUNTRY="Brazil" then response.write (" SELECTED ")%>>
								Brazil
							</option>
							<option <% if COUNTRY="British Indian Ocean Territory" then response.write (" SELECTED ")%>>
								British Indian Ocean Territory
							</option>
							<option <% if COUNTRY="Brunei Darussalam" then response.write (" SELECTED ")%>>
								Brunei Darussalam
							</option>
							<option <% if COUNTRY="Bulgaria" then response.write (" SELECTED ")%>>
								Bulgaria
							</option>
							<option <% if COUNTRY="Burkina Faso" then response.write (" SELECTED ")%>>
								Burkina Faso
							</option>
							<option <% if COUNTRY="Burma" then response.write (" SELECTED ")%>>
								Burma
							</option>
							<option <% if COUNTRY="Burundi" then response.write (" SELECTED ")%>>
								Burundi
							</option>
							<option <% if COUNTRY="Cambodia" then response.write (" SELECTED ")%>>
								Cambodia
							</option>
							<option <% if COUNTRY="Cameroon" then response.write (" SELECTED ")%>>
								Cameroon
							</option>
							<option <% if COUNTRY="Canada" then response.write (" SELECTED ")%>>
								Canada
							</option>
							<option <% if COUNTRY="Cape Verde" then response.write (" SELECTED ")%>>
								Cape Verde
							</option>
							<option <% if COUNTRY="Cayman Islands" then response.write (" SELECTED ")%>>
								Cayman Islands
							</option>
							<option <% if COUNTRY="Central African Republic" then response.write (" SELECTED ")%>>
								Central African Republic
							</option>
							<option <% if COUNTRY="Chad" then response.write (" SELECTED ")%>>
								Chad
							</option>
							<option <% if COUNTRY="Chile" then response.write (" SELECTED ")%>>
								Chile
							</option>
							<option <% if COUNTRY="China" then response.write (" SELECTED ")%>>
								China
							</option>
							<option <% if COUNTRY="Christmas Island" then response.write (" SELECTED ")%>>
								Christmas Island
							</option>
							<option <% if COUNTRY="Cocos (Keeling) Islands" then response.write (" SELECTED ")%>>
								Cocos (Keeling) Islands
							</option>
							<option <% if COUNTRY="Colombia" then response.write (" SELECTED ")%>>
								Colombia
							</option>
							<option <% if COUNTRY="Comoros" then response.write (" SELECTED ")%>>
								Comoros
							</option>
							<option <% if COUNTRY="Congo" then response.write (" SELECTED ")%>>
								Congo
							</option>
							<option <% if COUNTRY="Congo, the Democratic Republic of the" then response.write (" SELECTED ")%>>
								Congo, the Democratic Republic of the
							</option>
							<option <% if COUNTRY="Cook Islands" then response.write (" SELECTED ")%>>
								Cook Islands
							</option>
							<option <% if COUNTRY="Costa Rica" then response.write (" SELECTED ")%>>
								Costa Rica
							</option>
							<option <% if COUNTRY="Cote d'Ivoire" then response.write (" SELECTED ")%>>
								Cote d'Ivoire
							</option>
							<option <% if COUNTRY="Croatia" then response.write (" SELECTED ")%>>
								Croatia
							</option>
							<option <% if COUNTRY="Cyprus" then response.write (" SELECTED ")%>>
								Cyprus
							</option>
							<option <% if COUNTRY="Czech Republic" then response.write (" SELECTED ")%>>
								Czech Republic
							</option>
							<option <% if COUNTRY="Denmark" then response.write (" SELECTED ")%>>
								Denmark
							</option>
							<option <% if COUNTRY="Djibouti" then response.write (" SELECTED ")%>>
								Djibouti
							</option>
							<option <% if COUNTRY="Dominica" then response.write (" SELECTED ")%>>
								Dominica
							</option>
							<option <% if COUNTRY="Dominican Republic" then response.write (" SELECTED ")%>>
								Dominican Republic
							</option>
							<option <% if COUNTRY="East Timor" then response.write (" SELECTED ")%>>
								East Timor
							</option>
							<option <% if COUNTRY="Ecuador" then response.write (" SELECTED ")%>>
								Ecuador
							</option>
							<option <% if COUNTRY="Egypt" then response.write (" SELECTED ")%>>
								Egypt
							</option>
							<option <% if COUNTRY="El Salvador" then response.write (" SELECTED ")%>>
								El Salvador
							</option>
							<option <% if COUNTRY="England" then response.write (" SELECTED ")%>>
								England
							</option>
							<option <% if COUNTRY="Equatorial Guinea" then response.write (" SELECTED ")%>>
								Equatorial Guinea
							</option>
							<option <% if COUNTRY="Eritrea" then response.write (" SELECTED ")%>>
								Eritrea
							</option>
							<option <% if COUNTRY="Espana" then response.write (" SELECTED ")%>>
								Espana
							</option>
							<option <% if COUNTRY="Estonia" then response.write (" SELECTED ")%>>
								Estonia
							</option>
							<option <% if COUNTRY="Ethiopia" then response.write (" SELECTED ")%>>
								Ethiopia
							</option>
							<option <% if COUNTRY="Falkland Islands" then response.write (" SELECTED ")%>>
								Falkland Islands
							</option>
							<option <% if COUNTRY="Faroe Islands" then response.write (" SELECTED ")%>>
								Faroe Islands
							</option>
							<option <% if COUNTRY="Fiji" then response.write (" SELECTED ")%>>
								Fiji
							</option>
							<option <% if COUNTRY="Finland" then response.write (" SELECTED ")%>>
								Finland
							</option>
							<option <% if COUNTRY="France" then response.write (" SELECTED ")%>>
								France
							</option>
							<option <% if COUNTRY="French Guiana" then response.write (" SELECTED ")%>>
								French Guiana
							</option>
							<option <% if COUNTRY="French Polynesia" then response.write (" SELECTED ")%>>
								French Polynesia
							</option>
							<option <% if COUNTRY="French Southern Territories" then response.write (" SELECTED ")%>>
								French Southern Territories
							</option>
							<option <% if COUNTRY="Gabon" then response.write (" SELECTED ")%>>
								Gabon
							</option>
							<option <% if COUNTRY="Gambia" then response.write (" SELECTED ")%>>
								Gambia
							</option>
							<option <% if COUNTRY="Georgia" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if COUNTRY="Germany" then response.write (" SELECTED ")%>>
								Germany
							</option>
							<option <% if COUNTRY="Ghana" then response.write (" SELECTED ")%>>
								Ghana
							</option>
							<option <% if COUNTRY="Gibraltar" then response.write (" SELECTED ")%>>
								Gibraltar
							</option>
							<option <% if COUNTRY="Great Britain" then response.write (" SELECTED ")%>>
								Great Britain
							</option>
							<option <% if COUNTRY="Greece" then response.write (" SELECTED ")%>>
								Greece
							</option>
							<option <% if COUNTRY="Greenland" then response.write (" SELECTED ")%>>
								Greenland
							</option>
							<option <% if COUNTRY="Grenada" then response.write (" SELECTED ")%>>
								Grenada
							</option>
							<option <% if COUNTRY="Guadeloupe" then response.write (" SELECTED ")%>>
								Guadeloupe
							</option>
							<option <% if COUNTRY="Guam" then response.write (" SELECTED ")%>>
								Guam
							</option>
							<option <% if COUNTRY="Guatemala" then response.write (" SELECTED ")%>>
								Guatemala
							</option>
							<option <% if COUNTRY="Guinea" then response.write (" SELECTED ")%>>
								Guinea
							</option>
							<option <% if COUNTRY="Guinea-Bissau" then response.write (" SELECTED ")%>>
								Guinea-Bissau
							</option>
							<option <% if COUNTRY="Guyana" then response.write (" SELECTED ")%>>
								Guyana
							</option>
							<option <% if COUNTRY="Haiti" then response.write (" SELECTED ")%>>
								Haiti
							</option>
							<option <% if COUNTRY="Heard and Mc Donald Islands" then response.write (" SELECTED ")%>>
								Heard and Mc Donald Islands
							</option>
							<option <% if COUNTRY="Honduras" then response.write (" SELECTED ")%>>
								Honduras
							</option>
							<option <% if COUNTRY="Hong Kong" then response.write (" SELECTED ")%>>
								Hong Kong
							</option>
							<option <% if COUNTRY="Hungary" then response.write (" SELECTED ")%>>
								Hungary
							</option>
							<option <% if COUNTRY="Iceland" then response.write (" SELECTED ")%>>
								Iceland
							</option>
							<option <% if COUNTRY="India" then response.write (" SELECTED ")%>>
								India
							</option>
							<option <% if COUNTRY="Indonesia" then response.write (" SELECTED ")%>>
								Indonesia
							</option>
							<option <% if COUNTRY="Ireland" then response.write (" SELECTED ")%>>
								Ireland
							</option>
							<option <% if COUNTRY="Israel" then response.write (" SELECTED ")%>>
								Israel
							</option>
							<option <% if COUNTRY="Italy" then response.write (" SELECTED ")%>>
								Italy
							</option>
							<option <% if COUNTRY="Jamaica" then response.write (" SELECTED ")%>>
								Jamaica
							</option>
							<option <% if COUNTRY="Japan" then response.write (" SELECTED ")%>>
								Japan
							</option>
							<option <% if COUNTRY="Jordan" then response.write (" SELECTED ")%>>
								Jordan
							</option>
							<option <% if COUNTRY="Kazakhstan" then response.write (" SELECTED ")%>>
								Kazakhstan
							</option>
							<option <% if COUNTRY="Kenya" then response.write (" SELECTED ")%>>
								Kenya
							</option>
							<option <% if COUNTRY="Kiribati" then response.write (" SELECTED ")%>>
								Kiribati
							</option>
							<option <% if COUNTRY="Korea (North)" then response.write (" SELECTED ")%>>
								Korea(North)
							</option>
							<option <% if COUNTRY="Korea, Republic of" then response.write (" SELECTED ")%>>
								Korea, Republic of
							</option>
							<option <% if COUNTRY="Korea (South)" then response.write (" SELECTED ")%>>
								Korea (South)
							</option>
							<option <% if COUNTRY="Kuwait" then response.write (" SELECTED ")%>>
								Kuwait
							</option>
							<option <% if COUNTRY="Kyrgyzstan" then response.write (" SELECTED ")%>>
								Kyrgyzstan
							</option>
							<option <% if COUNTRY="Lao People's Democratic Republic" then response.write (" SELECTED ")%>>
								Lao People's Democratic Republic
							</option>
							<option <% if COUNTRY="Latvia" then response.write (" SELECTED ")%>>
								Latvia
							</option>
							<option <% if COUNTRY="Lebanon" then response.write (" SELECTED ")%>>
								Lebanon
							</option>
							<option <% if COUNTRY="Lesotho" then response.write (" SELECTED ")%>>
								Lesotho
							</option>
							<option <% if COUNTRY="Liberia" then response.write (" SELECTED ")%>>
								Liberia
							</option>
							<option <% if COUNTRY="Liechtenstein" then response.write (" SELECTED ")%>>
								Liechtenstein
							</option>
							<option <% if COUNTRY="Lithuania" then response.write (" SELECTED ")%>>
								Lithuania
							</option>
							<option <% if COUNTRY="Luxembourg" then response.write (" SELECTED ")%>>
								Luxembourg
							</option>
							<option <% if COUNTRY="Macau" then response.write (" SELECTED ")%>>
								Macau
							</option>
							<option <% if COUNTRY="Macedonia" then response.write (" SELECTED ")%>>
								Macedonia
							</option>
							<option <% if COUNTRY="Madagascar" then response.write (" SELECTED ")%>>
								Madagascar
							</option>
							<option <% if COUNTRY="Malawi" then response.write (" SELECTED ")%>>
								Malawi
							</option>
							<option <% if COUNTRY="Malaysia" then response.write (" SELECTED ")%>>
								Malaysia
							</option>
							<option <% if COUNTRY="Maldives" then response.write (" SELECTED ")%>>
								Maldives
							</option>
							<option <% if COUNTRY="Mali" then response.write (" SELECTED ")%>>
								Mali
							</option>
							<option <% if COUNTRY="Malta" then response.write (" SELECTED ")%>>
								Malta
							</option>
							<option <% if COUNTRY="Marshall Islands" then response.write (" SELECTED ")%>>
								Marshall Islands
							</option>
							<option <% if COUNTRY="Martinique" then response.write (" SELECTED ")%>>
								Martinique
							</option>
							<option <% if COUNTRY="Mauritania" then response.write (" SELECTED ")%>>
								Mauritania
							</option>
							<option <% if COUNTRY="Mauritius" then response.write (" SELECTED ")%>>
								Mauritius
							</option>
							<option <% if COUNTRY="Mayotte" then response.write (" SELECTED ")%>>
								Mayotte
							</option>
							<option <% if COUNTRY="Mexico" then response.write (" SELECTED ")%>>
								Mexico
							</option>
							<option <% if COUNTRY="Micronesia, Federated States of" then response.write (" SELECTED ")%>>
								Micronesia, Federated States of
							</option>
							<option <% if COUNTRY="Moldova, Republic of" then response.write (" SELECTED ")%>>
								Moldova, Republic of
							</option>
							<option <% if COUNTRY="Monaco" then response.write (" SELECTED ")%>>
								Monaco
							</option>
							<option <% if COUNTRY="Mongolia" then response.write (" SELECTED ")%>>
								Mongolia
							</option>
							<option <% if COUNTRY="Montserrat" then response.write (" SELECTED ")%>>
								Montserrat
							</option>
							<option <% if COUNTRY="Morocco" then response.write (" SELECTED ")%>>
								Morocco
							</option>
							<option <% if COUNTRY="Mozambique" then response.write (" SELECTED ")%>>
								Mozambique
							</option>
							<option <% if COUNTRY="Myanmar" then response.write (" SELECTED ")%>>
								Myanmar
							</option>
							<option <% if COUNTRY="Namibia" then response.write (" SELECTED ")%>>
								Namibia
							</option>
							<option <% if COUNTRY="Nauru" then response.write (" SELECTED ")%>>
								Nauru
							</option>
							<option <% if COUNTRY="Nepal" then response.write (" SELECTED ")%>>
								Nepal
							</option>
							<option <% if COUNTRY="Netherlands" then response.write (" SELECTED ")%>>
								Netherlands
							</option>
							<option <% if COUNTRY="Netherlands Antilles" then response.write (" SELECTED ")%>>
								NetherlandsAntilles
							</option>
							<option <% if COUNTRY="New Caledonia" then response.write (" SELECTED ")%>>
								New Caledonia
							</option>
							<option <% if COUNTRY="New Zealand" then response.write (" SELECTED ")%>>
								New Zealand
							</option>
							<option <% if COUNTRY="Nicaragua" then response.write (" SELECTED ")%>>
								Nicaragua
							</option>
							<option <% if COUNTRY="Niger" then response.write (" SELECTED ")%>>
								Niger
							</option>
							<option <% if COUNTRY="Nigeria" then response.write (" SELECTED ")%>>
								Nigeria
							</option>
							<option <% if COUNTRY="Niue" then response.write (" SELECTED ")%>>
								Niue
							</option>
							<option <% if COUNTRY="Norfolk Island" then response.write (" SELECTED ")%>>
								Norfolk Island
							</option>
							<option <% if COUNTRY="Northern Ireland" then response.write (" SELECTED ")%>>
								Northern Ireland
							</option>
							<option <% if COUNTRY="Northern Mariana Islands" then response.write (" SELECTED ")%>>
								Northern Mariana Islands
							</option>
							<option <% if COUNTRY="Norway" then response.write (" SELECTED ")%>>
								Norway
							</option>
							<option <% if COUNTRY="Oman" then response.write (" SELECTED ")%>>
								Oman
							</option>
							<option <% if COUNTRY="Pakistan" then response.write (" SELECTED ")%>>
								Pakistan
							</option>
							<option <% if COUNTRY="Palau" then response.write (" SELECTED ")%>>
								Palau
							</option>
							<option <% if COUNTRY="Panama" then response.write (" SELECTED ")%>>
								Panama
							</option>
							<option <% if COUNTRY="Papua New Guinea" then response.write (" SELECTED ")%>>
								Papua New Guinea
							</option>
							<option <% if COUNTRY="Paraguay" then response.write (" SELECTED ")%>>
								Paraguay
							</option>
							<option <% if COUNTRY="Peru" then response.write (" SELECTED ")%>>
								Peru
							</option>
							<option <% if COUNTRY="Philippines" then response.write (" SELECTED ")%>>
								Philippines
							</option>
							<option <% if COUNTRY="Pitcairn" then response.write (" SELECTED ")%>>
								Pitcairn
							</option>
							<option <% if COUNTRY="Poland" then response.write (" SELECTED ")%>>
								Poland
							</option>
							<option <% if COUNTRY="Portugal" then response.write (" SELECTED ")%>>
								Portugal
							</option>
							<option <% if COUNTRY="Puerto Rico" then response.write (" SELECTED ")%>>
								Puerto Rico
							</option>
							<option <% if COUNTRY="Qatar" then response.write (" SELECTED ")%>>
								Qatar
							</option>
							<option <% if COUNTRY="Reunion" then response.write (" SELECTED ")%>>
								Reunion
							</option>
							<option <% if COUNTRY="Romania" then response.write (" SELECTED ")%>>
								Romania
							</option>
							<option <% if COUNTRY="Russia" then response.write (" SELECTED ")%>>
								Russia
							</option>
							<option <% if COUNTRY="Russian Federation" then response.write (" SELECTED ")%>>
								Russian Federation
							</option>
							<option <% if COUNTRY="Rwanda" then response.write (" SELECTED ")%>>
								Rwanda
							</option>
							<option <% if COUNTRY="Saint Kitts and Nevis" then response.write (" SELECTED ")%>>
								Saint Kitts and Nevis
							</option>
							<option <% if COUNTRY="Saint Lucia" then response.write (" SELECTED ")%>>
								Saint Lucia
							</option>
							<option <% if COUNTRY="Saint Vincent and the Grenadines" then response.write (" SELECTED ")%>>
								Saint Vincent and the Grenadines
							</option>
							<option <% if COUNTRY="Samoa (Independent)" then response.write (" SELECTED ")%>>
								Samoa (Independent)
							</option>
							<option <% if COUNTRY="San Marino" then response.write (" SELECTED ")%>>
								San Marino
							</option>
							<option <% if COUNTRY="Sao Tome and Principe" then response.write (" SELECTED ")%>>
								Sao Tome and Principe
							</option>
							<option <% if COUNTRY="Saudi Arabia" then response.write (" SELECTED ")%>>
								Saudi Arabia
							</option>
							<option <% if COUNTRY="Scotland" then response.write (" SELECTED ")%>>
								Scotland
							</option>
							<option <% if COUNTRY="Senegal" then response.write (" SELECTED ")%>>
								Senegal
							</option>
							<option <% if COUNTRY="Seychelles" then response.write (" SELECTED ")%>>
								Seychelles
							</option>
							<option <% if COUNTRY="Sierra Leone" then response.write (" SELECTED ")%>>
								SierraLeone
							</option>
							<option <% if COUNTRY="Singapore" then response.write (" SELECTED ")%>>
								Singapore
							</option>
							<option <% if COUNTRY="Slovakia" then response.write (" SELECTED ")%>>
								Slovakia
							</option>
							<option <% if COUNTRY="Slovenia" then response.write (" SELECTED ")%>>
								Slovenia
							</option>
							<option <% if COUNTRY="Solomon Islands" then response.write (" SELECTED ")%>>
								Solomon Islands
							</option>
							<option <% if COUNTRY="Somalia" then response.write (" SELECTED ")%>>
								Somalia
							</option>
							<option <% if COUNTRY="South Africa" then response.write (" SELECTED ")%>>
								South Africa
							</option>
							<option <% if COUNTRY="South Georgia and the South Sandwich Islands" then response.write (" SELECTED ")%>>
								South Georgia and the South Sandwich Islands
							</option>
							<option <% if COUNTRY="South Korea" then response.write (" SELECTED ")%>>
								South Korea
							</option>
							<option <% if COUNTRY="Spain" then response.write (" SELECTED ")%>>
								Spain
							</option>
							<option <% if COUNTRY="Sri Lanka" then response.write (" SELECTED ")%>>
								Sri Lanka
							</option>
							<option <% if COUNTRY="St. Helena" then response.write (" SELECTED ")%>>
								St. Helena
							</option>
							<option <% if COUNTRY="St. Pierre and Miquelon" then response.write (" SELECTED ")%>>
								St. Pierre and Miquelon
							</option>
							<option <% if COUNTRY="Suriname" then response.write (" SELECTED ")%>>
								Suriname
							</option>
							<option <% if COUNTRY="Svalbard and Jan Mayen Islands" then response.write (" SELECTED ")%>>
								Svalbard and Jan Mayen Islands
							</option>
							<option <% if COUNTRY="Swaziland" then response.write (" SELECTED ")%>>
								Swaziland
							</option>
							<option <% if COUNTRY="Sweden" then response.write (" SELECTED ")%>>
								Sweden
							</option>
							<option <% if COUNTRY="Switzerland" then response.write (" SELECTED ")%>>
								Switzerland
							</option>
							<option <% if COUNTRY="Taiwan" then response.write (" SELECTED ")%>>
								Taiwan
							</option>
							<option <% if COUNTRY="Tajikistan" then response.write (" SELECTED ")%>>
								Tajikistan
							</option>
							<option <% if COUNTRY="Tanzania" then response.write (" SELECTED ")%>>
								Tanzania
							</option>
							<option <% if COUNTRY="Thailand" then response.write (" SELECTED ")%>>
								Thailand
							</option>
							<option <% if COUNTRY="Togo" then response.write (" SELECTED ")%>>
								Togo
							</option>
							<option <% if COUNTRY="Tokelau" then response.write (" SELECTED ")%>>
								Tokelau
							</option>
							<option <% if COUNTRY="Tonga" then response.write (" SELECTED ")%>>
								Tonga
							</option>
							<option <% if COUNTRY="Trinidad" then response.write (" SELECTED ")%>>
								Trinidad
							</option>
							<option <% if COUNTRY="Trinidad and Tobago" then response.write (" SELECTED ")%>>
								Trinidad and Tobago
							</option>
							<option <% if COUNTRY="Tunisia" then response.write (" SELECTED ")%>>
								Tunisia
							</option>
							<option <% if COUNTRY="Turkey" then response.write (" SELECTED ")%>>
								Turkey
							</option>
							<option <% if COUNTRY="Turkmenistan" then response.write (" SELECTED ")%>>
								Turkmenistan
							</option>
							<option <% if COUNTRY="Turks and Caicos Islands" then response.write (" SELECTED ")%>>
								Turks and Caicos Islands
							</option>
							<option <% if COUNTRY="Tuvalu" then response.write (" SELECTED ")%>>
								Tuvalu
							</option>
							<option <% if COUNTRY="Uganda" then response.write (" SELECTED ")%>>
								Uganda
							</option>
							<option <% if COUNTRY="Ukraine" then response.write (" SELECTED ")%>>
								Ukraine
							</option>
							<option <% if COUNTRY="United Arab Emirates" then response.write (" SELECTED ")%>>
								United Arab Emirates
							</option>
							<option <% if COUNTRY="United Kingdom" then response.write (" SELECTED ")%>>
								United Kingdom
							</option>
							<option <% if COUNTRY="United States" then response.write (" SELECTED ")%>>
								United States
							</option>
							<option <% if COUNTRY="United States Minor Outlying Islands" then response.write (" SELECTED ")%>>
								United States Minor Outlying Islands
							</option>
							<option <% if COUNTRY="Uruguay" then response.write (" SELECTED ")%>>
								Uruguay
							</option>
							<option <% if COUNTRY="Uzbekistan" then response.write (" SELECTED ")%>>
								Uzbekistan
							</option>
							<option <% if COUNTRY="Vanuatu" then response.write (" SELECTED ")%>>
								Vanuatu
							</option>
							<option <% if COUNTRY="Vatican City State (Holy See)" then response.write (" SELECTED ")%>>
								Vatican City State (Holy See)
							</option>
							<option <% if COUNTRY="Venezuela" then response.write (" SELECTED ")%>>
								Venezuela
							</option>
							<option <% if COUNTRY="Viet Nam" then response.write (" SELECTED ")%>>
								Viet Nam
							</option>
							<option <% if COUNTRY="Virgin Islands (British)" then response.write (" SELECTED ")%>>
								Virgin Islands (British)
							</option>
							<option <% if COUNTRY="Virgin Islands (U.S.)" then response.write (" SELECTED ")%>>
								Virgin Islands (U.S.)
							</option>
							<option <% if COUNTRY="Wales" then response.write (" SELECTED ")%>>
								Wales
							</option>
							<option <% if COUNTRY="Wallis and Futuna Islands" then response.write (" SELECTED ")%>>
								Wallis and Futuna Islands
							</option>
							<option <% if COUNTRY="Western Sahara" then response.write (" SELECTED ")%>>
								Western Sahara
							</option>
							<option <% if COUNTRY="Yemen" then response.write (" SELECTED ")%>>
								Yemen
							</option>
							<option <% if COUNTRY="Zambia" then response.write (" SELECTED ")%>>
								Zambia
							</option>
							<option <% if COUNTRY="Zimbabwe" then response.write (" SELECTED ")%>>
								Zimbabwe
							</option>
						</select>
					</td>
					<td>
						<div align="right">
							Cell</div>
					</td>
					<td><input type=text value="<%=CELL_PHONE%>" name="CELL_PHONE2" size=20 maxlength=20 tabindex="13">
</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Email</div>
					</td>
					<td width=85%>
					  <input type=text value="<%=EMAIL%>" name="EMAIL" size=50 maxlength=50 tabindex="12">
					</td>
					<td width=85%>&nbsp;
					</td>
					<td width=85%>&nbsp;					</td>
				</tr>
				<tr>
					<td colspan="4">
						<div align="left">
							<strong>
<input type="checkbox" name="SHARE_INFO" value="No" tabindex="14" <% if SHARE_INFO="No" then response.write (" CHECKED ")%> >								
<font color="#FF0000">
					Do Not Share Info
				</font>
							</strong>
						</div>
					</td>
				</tr>
			</table>
			<table border=0 width=100% cellspacing="0">
<!-- BEGIN GRADE LEVEL INFORMATION-->
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<div align="left">
							<strong>
								Grade Level Information
							</strong>
						</div>
					</td>
					<td width=22%>
					</td>
					<td width=20%>
					</td>
					<td width=12%>
					</td>
					<td width=16%>
					</td>
				</tr>
				<tr>
					<td width=9%>
					</td>
					<td width=21% valign="top">
						<input type="checkbox" name="GRADE_LEVELS" value="1" tabindex="15" <% if instr (GRADE_LEVELS,"1") then response.write (" checked ")%> >
						Elementary<br>
					</td>
					<td width=22% valign="top">
						<input type="checkbox" name="GRADE_LEVELS" value="2" tabindex="15" <% if instr (GRADE_LEVELS,"2") then response.write (" checked ")%> >
						Junior High School
					</td>
					<td width=20% valign="top">
						<input type="checkbox" name="GRADE_LEVELS" value="3" tabindex="17" <% if instr (GRADE_LEVELS,"3") then response.write (" checked ")%> >
						High School<br>
					</td>
					<td colspan="2" valign="top">
						<div align="left">
						  <input type="checkbox" name="GRADE_LEVELS" value="4" tabindex="18" <% if instr (GRADE_LEVELS,"4") then response.write (" checked ")%> >
						College
				    	</div>
					</td>
				</tr>
<!-- BEGIN POSITION INFORMATION-->
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<strong>
							Position Information
						</strong>
					</td>
					<td width=22%>
					</td>
					<td width=20%>
					</td>
					<td width=12% bgcolor="#AAD5FF">
					</td>
					<td width=16%>
					</td>
				</tr>
				<tr>
					<td width=9%>
					</td>
					<td width=21% valign="top">
						<input type="checkbox" name="POSITION" value="B" <% if instr (POSITION,"B") then response.write (" checked ")%> tabindex="19" >
						Band Director<br>
						<input type="checkbox" name="POSITION" value="O" tabindex="22" <% if instr (POSITION,"O") then response.write (" checked ")%> >
					  Orchestra Director
						<br>
						<input type="checkbox" name="POSITION" value="A" tabindex="24" <% if instr (POSITION,"A") then response.write (" checked ")%> >
						Administrator<br>
						<input type="checkbox" name="POSITION" value="J" tabindex="27" <% if instr (POSITION,"J") then response.write (" checked ")%> >
						Jazz
					</td>
					<td width=22% valign="top">
						<input type="checkbox" name="POSITION" value="P" tabindex="20" <% if instr (POSITION,"P") then response.write (" checked ")%> >
						Private Teacher<br>
						<input type="checkbox" name="POSITION" value="C" tabindex="22" <% if instr (POSITION,"C") then response.write (" checked ")%> >
					  Composer/Arranger
						<br>
						<input type="checkbox" name="POSITION" value="L" tabindex="25" <% if instr (POSITION,"L") then response.write (" checked ")%> >
					  Commercial
						<br>
						<input type="checkbox" name="POSITION" value="S" tabindex="28" <% if instr (POSITION,"S") then response.write (" checked ")%> >
						College Student
					</td>
				  	<td width=20% valign="top">						
						<input type="checkbox" name="POSITION" value="R" tabindex="21" <% if instr (POSITION,"R") then response.write (" checked ")%> >
						Retired<br>
						<input type="checkbox" name="POSITION" value="M" tabindex="23" <% if instr (POSITION,"M") then response.write (" checked ")%> >
						Military<br>
                        <input type="checkbox" name="POSITION" value="CBO" tabindex="26" <% if instr (POSITION,"CBO") then response.write (" checked ")%> >					  
                      Community Band/Orchestra
					  <br>
					  	<input type="checkbox" name="POSITION" value="X" tabindex="29" <% if instr (POSITION,"X") then response.write (" checked ")%> >
						Other
					</td>
					<td colspan="2" valign="top">
					</td>
				</tr>
				<tr>
					<td colspan="6">&nbsp;
					</td>
				</tr>
				<tr bgcolor="#AAD5FF">
					<td colspan="6" valign="bottom">
						<strong>
							Hotel While At Conference
						</strong>
						&nbsp;&nbsp;
						<select name="HOTEL" tabindex="30">
							<option >
								&nbsp;
							</option>
							<option <% if HOTEL="Blackstone Hotel Downtown" then response.write ("SELECTED ")%>>
								Blackstone
							</option>		
							<option <% if HOTEL="Embassy Suites Chicago" then response.write ("SELECTED ")%>>
								Embassy Suites Chicago
							</option>
							<option <% if HOTEL="Essex Inn" then response.write ("SELECTED ")%>>
								Essex Inn
							</option>
							<option <% if HOTEL="Hilton Chicago" then response.write ("SELECTED ")%>>
								Hilton Chicago
							</option>
							<option <% if HOTEL="Hyatt Regency Downtown" then response.write ("SELECTED ")%>>
								Hyatt Regency Downtown
							</option>
							<option <% if HOTEL="Hyatt Regency McCormick Place" then response.write ("SELECTED ")%>>
								Hyatt Regency McCormick Place
							</option>
							<option <% if HOTEL="Palmer House" then response.write ("SELECTED ")%>>
								Palmer House
							</option>
							<option <% if HOTEL="Sheraton Chicago Hotel" then response.write ("SELECTED ")%>>
								Sheraton Chicago Hotel
							</option>					
							<option <% if HOTEL="Other" then response.write ("SELECTED ")%>>
								Other
							</option>
						</select>
					</td>
				</tr>
				<tr>
					<td colspan="6">&nbsp;
					</td>
				</tr>
				<tr>
					<td bgcolor="#AAD5FF" colspan="6">
						<strong>
							Emergency Contact Number
						</strong>
						&nbsp;&nbsp;
						<input type=text value="<%=EMERGENCY_PHONE%>" name="EMERGENCY_PHONE" size=20 maxlength=20 tabindex="31">
					</td>
				</tr>
				<tr>
					<td colspan="6">&nbsp;
					</td>
				</tr>
<!-- BEGIN BADGE INFO -->
				<tr bgcolor="#AAD5FF">
					<td width=9%>
						<strong>
							Badge Type
						</strong>
					</td>
					<td colspan=2>
						<strong>
							<input type="checkbox" name="BADGE_COMP" value="Y" <% if BADGE_COMP="Y" then response.write (" CHECKED ")%> onClick="DisplayAmountDue(this.value)" tabindex="32">
							Make Complimentary
						</strong>
					</td>
					<td width=20%>
					</td>
					<td width=12%>
					</td>
					<td width=16%>
					</td>
				</tr>
				<tr>
					<td width=9% valign="top">
						<a href="attendee_badge_info.asp" tabindex="33">
							Badge Activity
						</a>
					</td>
					<td width=21%>
						<input type="checkbox" name="BADGE_TYPE" value="D" <% if instr (BADGE_TYPE,"D") then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="33">
						Adult/Director
						<br>
						<input type="checkbox" name="BADGE_TYPE" value="S" <% if BADGE_TYPE = "S" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="36">
						College Student</td>
					<td width=22%>
						<input type="checkbox" name="BADGE_TYPE" value="L" <% if BADGE_TYPE = "L" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="34">
						Commercial<br>
						<input type="checkbox" name="BADGE_TYPE" value="C" <% if BADGE_TYPE = "C" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="37">
						Clinician</td>
					<td width=20%>
					  <input type="checkbox" name="BADGE_TYPE" value="G" <% if BADGE_TYPE = "G" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="35">
					  Guest
					  <br>
						<input type="checkbox" name="BADGE_TYPE" value="F" <% if BADGE_TYPE = "F" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="36">
						Staff</td>
					<td colspan="2" align="left" valign="top">
						<input type="checkbox" name="BADGE_TYPE" value="BS" <% if BADGE_TYPE = "BS" then response.write (" checked ")%> onClick="DisplayAmountDue(this.value)" tabindex="37"> 
					  Board Spouse</td>
				</tr>
			</table>
<!-- BEGIN PAYMENT INFORMATION-->
			<table border=0 width=100% cellspacing="0">
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<strong>
							Payment Information
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
						<input type=text value="<%=iCnt%>" name="ADDL_PROG_COUNT" size="2" onKeyUp="CalcProgCost()" onreadystatechange="CalcProgCost()" tabindex="39">
						&nbsp;Program Fee
					</td>
					<td width=8%>
						<input type=text value="<%=(iCnt * 20)%>" name="ADDL_PROG_COST" size=10 readonly="true">
					</td>
					<td width=16%>
					</td>
					<td width=48%>
					</td>
				</tr>
				<tr>
					<td width=12% style="height: 24px">
					</td>
					<td width=16% style="height: 24px">
						<div  align="right">
							Registration Fee
						</div>
					</td>
					<td width=8% style="height: 24px">
						<input type=text value="<%=AMOUNT_DUE%>" name="AMOUNT_DUE" size=10 readonly="true">
					</td>
					<td width=16% style="height: 24px">
					</td>
					<td width=48% style="height: 24px">
					</td>
				</tr>
				<% iTotal = (iCnt * 20) + AMOUNT_DUE %>
				<tr>
					<td width=12% style="height: 24px">
						<a href="attendee_payment_info.asp" tabindex="40">
							Payment Activity
						</a>
					</td>
					<td width=16% style="height: 24px">
						<div  align="right">
							Total Cost For Attendee
						</div>
					</td>
					<td width=8% style="height: 24px">
						<input type=text value="<%=iTotalPayment%>" name="TOTAL_AMOUNT_DUE" size=10 readonly="true">
					</td>
					<td width=16% style="height: 24px">
					</td>
					<td width=48% style="height: 24px">
					</td>
				</tr>
				<tr>
					<td width=12%>
					</td>
					<td width=16%>
						<div align="right">
							Payment
						</div>
					</td>
					<td width=8% >
						<input type=text value="<%=iTotalPayment%>" name="PAYMENT" size=10 tabindex="41">
					</td>
					<td width=16% rowspan="2" valign="top">
						<input type="checkbox" name="PAYMENT_TYPE" value="CC" <% if instr (sPaymentType,"CC") then response.write (" checked ")%> tabindex="42">
						Credit Card <br>
						<input type="checkbox" name="PAYMENT_TYPE" value="CH" <% if instr (sPaymentType,"CH") then response.write (" checked ")%> tabindex="43">
						Check <br>
						<input type="checkbox" name="PAYMENT_TYPE" value="CA" <% if instr (sPaymentType,"CA") then response.write (" checked ")%> tabindex="44">
						Cash/Trav Check <br>
						<input type="checkbox" name="PAYMENT_TYPE" value="NC" <% if instr (sPaymentType,"NC") then response.write (" checked ")%> onClick="NoChargeClicked()" tabindex="44">
						No Charge
					</td>
					<td width=48%>
					</td>
				</tr>
				<tr>
					<td width=12%>
					</td>
					<td width=16%>
						<div align="right">
							Initials
						</div>
					</td>
					<td width=8%>
                      <input type=text value="<%=sDisplayInitials%>" name="BADGE_INITIALS" size=5 maxlength="5" tabindex="45">
                    </td>
					<td width=16%>
					</td>
					<td width=48%>
					</td>
				</tr>
				<% iTotalDue = iTotal - iTotalPayment %>
				<tr>
					<td width=12%>
					</td>
					<td width=14%>
						<div align="right">
							Total Amount Due
						</div>
					</td>
					<td width=8%>
						<input type=text value="<%=iTotalDue %>" name="TO_BE_PAID" size=10 readonly="true">
					</td>
					<td width=16%>
						<input type="hidden" value="<%=iTotalDue %>" name="TO_BE_PAID_INITIAL" tabindex="">
					</td>
					<td width=48%>
					</td>
				</tr>
			</table>
			<table border=0 width="100%">
				<tr>
					<td width="40%" align="left">
						Comment&nbsp;
						<input type=text value="<%=COMMENTS %>" name="COMMENTS" size=50 maxlength="50" style="width: 309px" tabindex="46">
					</td>
					<td width=20%>
						<div align="center">
							<strong>
								<input type="checkbox" name="PRINT_BADGE" tabindex="47" value="PRINT_BADGE" 
			<% if Session("ATTENDEE_LIST") = "" or sPreregistered = "Y" or Session("DUPLICATE") <> "" then response.write (" checked ")%> >
								Print Badge
							</strong>
						</div>
					</td>
					<td width="40%">
					</td>
				</tr>
				<tr>
					<td colspan=3>
						<div align="center">
                          <input type="submit" name="SUBMIT" value="Submit" tabindex="48">							
                          <input type="reset" name="RESET" value="Reset" tabindex="49">
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="3">
						<div align="center">
							<a href="select_attendee.asp" tabindex="50">
								Return to Select Attendees
							</a>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
