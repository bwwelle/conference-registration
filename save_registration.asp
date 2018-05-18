<% @Language=VBScript %>
<% Response.Buffer="true"%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim conDB
	dim sessionFirst
	dim sessionLast
	dim sessionCity
	dim sessionState
	dim strSQL
	dim lngRecs
	dim rsRegID
	dim iRegID
	dim rsUpdReg
	dim rsBadgeID
	dim iBadgeID
	dim rsUpdBadge
	dim rsProgID
	dim rsBadgeActivity
	dim rsPaymentActivity
	dim arAttendeeField
	dim newRegID
	dim newBadgeID
	dim iAmt
	dim iAmtDue
	dim iTotAmtPd
	dim iAmtOut
	dim iserror
	dim iCnt
	dim sFIRST_NAME
	dim sLAST_NAME
	dim sADDRESS
	dim sCITY
	dim sSTATE
	dim sZIP
	dim sCOUNTRY
	dim sEMAIL
	dim sCELL_PHONE
	dim sEMERGENCY_PHONE
	dim sCOMMENTS
	dim cBADGE_TYPE
	dim cBADGE_COMP
	dim sBADGE_INITIALS
	dim cSHARE_INFO
	dim cGRADE_LEVELS
	dim cPOSITION
	dim sAMOUNT_DUE
	dim sPAYMENT
	dim cPAYMENT_TYPE
	dim cPRINT_BADGE
	dim sBADGE_PICKUP_NAME
	dim sPre
	dim sHOTEL
	dim sSCHOOL_COMPANY
	dim sADDL_PROG_COUNT
	dim sADDL_PROG_COST
	dim iRegPayment
	
		
	dim sBadgeString
	dim sReceiptString
	dim rs
	dim sPayType
	dim sAmount
	dim s75
	dim s25
	dim sZero
	dim sName
	dim sFrom
	dim sBadgeName
	dim sBadgeFrom
	dim btype
	dim sSchoolCompany
	dim sSC1
	dim sSC2
	
	sFIRST_NAME			= Trim(Request("FIRST_NAME"))
	sLAST_NAME			= Trim(Request("LAST_NAME"))
	sADDRESS			= Trim(Request("ADDRESS"))
	
	Session("ADDEDATTENDEE") = "true"
	
	If Trim(Request("CITY")) = "" then
		sCITY = ""
	Else
		sCITY = Trim(Request("CITY"))
		sCITY = Replace(sCITY,"'","''")
	End If
	
	If Trim(Request("STATES")) = "" then
		sSTATE = ""
	Else
		sSTATE = Trim(Request("STATES"))
	End If
	
	sZIP				= Trim(Request("ZIP"))
	sCOUNTRY			= Trim(Request("COUNTRY"))
	sEMAIL				= Trim(Request("EMAIL"))
	sCELL_PHONE			= Trim(Request("CELL_PHONE"))
	sEMERGENCY_PHONE	= Trim(Request("EMERGENCY_PHONE"))
	sCOMMENTS			= Trim(Request("COMMENTS"))
	cBADGE_TYPE			= Trim(Request("BADGE_TYPE"))
	cBADGE_COMP			= Trim(Request("BADGE_COMP"))
	sBADGE_INITIALS		= Trim(Request("BADGE_INITIALS"))
	cSHARE_INFO			= Trim(Request("SHARE_INFO"))
	cGRADE_LEVELS		= Trim(Request("GRADE_LEVELS"))
	cPOSITION 			= Trim(Request("POSITION"))
	sAMOUNT_DUE	      	= Trim(Request("AMOUNT_DUE"))
	sPAYMENT      		= Trim(Request("PAYMENT"))		'the total payment
	cPAYMENT_TYPE		= Trim(Request("PAYMENT_TYPE"))
	cPRINT_BADGE		= Trim(Request("PRINT_BADGE"))
	sBADGE_PICKUP_NAME	= Trim(Request("BADGE_PICKUP_NAME"))
	sTO_BE_PAID			= Trim(Request("TO_BE_PAID"))   'the total due before this payment
	sHOTEL 				= Trim(Request("HOTEL"))
	sSCHOOL_COMPANY		= Trim(Request("SCHOOL_COMPANY"))
	sADDL_PROG_COUNT	= Trim(Request("ADDL_PROG_COUNT")) 'programs if any
	sADDL_PROG_COST		= Trim(Request("ADDL_PROG_COST"))
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if no programs, set program count to 0.  if there are
' programs, subtract that payment from the total payment
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If sADDL_PROG_COUNT = "" Then
		sADDL_PROG_COUNT = "0"
	End If
	
	If sPAYMENT = "" Then
		sPAYMENT = "0"
	End If	
	
	iRegPayment = CInt(sPAYMENT) - (CInt(sADDL_PROG_COUNT) * 20)

	Session("BADGE_INITIALS") = sBADGE_INITIALS
	
	If sSTATE = "Alabama" Then sSTATE = "AL"
	If sSTATE = "Alaska" Then sSTATE = "AK"
	If sSTATE = "Arizona" Then sSTATE = "AZ"
	If sSTATE = "Arkansas" Then sSTATE = "AR"
	If sSTATE = "California" Then sSTATE = "CA"
	If sSTATE = "Colorado" Then sSTATE = "CO"
	If sSTATE = "Connecticut" Then sSTATE = "CT"
	If sSTATE = "Delaware" Then sSTATE = "DE"
	If sSTATE = "District of Columbia" Then sSTATE = "DC"
	If sSTATE = "Florida" Then sSTATE = "FL"
	If sSTATE = "Georgia" Then sSTATE = "GA"
	If sSTATE = "Guam" Then sSTATE = "GU"
	If sSTATE = "Hawaii" Then sSTATE = "HI"
	If sSTATE = "Idaho" Then sSTATE = "ID"
	If sSTATE = "Illinois" Then sSTATE = "IL"
	If sSTATE = "Indiana" Then sSTATE = "IN"
	If sSTATE = "Iowa" Then sSTATE = "IA"
	If sSTATE = "Kansas" Then sSTATE = "KS"
	If sSTATE = "Kentucky" Then sSTATE = "KY"
	If sSTATE = "Louisiana" Then sSTATE = "LA"
	If sSTATE = "Maine" Then sSTATE = "ME"
	If sSTATE = "Maryland" Then sSTATE = "MD"
	If sSTATE = "Massachusetts" Then sSTATE = "MA"
	If sSTATE = "Michigan" Then sSTATE = "MI"
	If sSTATE = "Minnesota" Then sSTATE = "MN"
	If sSTATE = "Mississippi" Then sSTATE = "MS"
	If sSTATE = "Missouri" Then sSTATE = "MO"
	If sSTATE = "Montana" Then sSTATE = "MT"
	If sSTATE = "Nebraska" Then sSTATE = "NE"
	If sSTATE = "Nevada" Then sSTATE = "NV"
	If sSTATE = "New Hampshire" Then sSTATE = "NH"
	If sSTATE = "New Jersey" Then sSTATE = "NJ"
	If sSTATE = "New Mexico" Then sSTATE = "NM"
	If sSTATE = "New York" Then sSTATE = "NY"
	If sSTATE = "North Carolina" Then sSTATE = "NC"
	If sSTATE = "North Dakota" Then sSTATE = "ND"
	If sSTATE = "Ohio" Then sSTATE = "OH"
	If sSTATE = "Oklahoma" Then sSTATE = "OK"
	If sSTATE = "Oregon" Then sSTATE = "OR"
	If sSTATE = "Pennsylvania" Then sSTATE = "PA"
	If sSTATE = "Puerto Rico" Then sSTATE = "PR"
	If sSTATE = "Rhode Island" Then sSTATE = "RI"
	If sSTATE = "South Carolina" Then sSTATE = "SC"
	If sSTATE = "South Dakota" Then sSTATE = "SD"
	If sSTATE = "Tennessee" Then sSTATE = "TN"
	If sSTATE = "Texas" Then sSTATE = "TX"
	If sSTATE = "Utah" Then sSTATE = "UT"
	If sSTATE = "Vermont" Then sSTATE = "VT"
	If sSTATE = "Virgin Islands" Then sSTATE = "VI"
	If sSTATE = "Virginia" Then sSTATE = "VA"
	If sSTATE = "Washington" Then sSTATE = "WA"
	If sSTATE = "West Virginia" Then sSTATE = "WV"
	If sSTATE = "Wisconsin" Then sSTATE = "WI"
	If sSTATE = "Wyoming" Then sSTATE = "WY"
	If sSTATE = "Alberta" Then sSTATE = "AB"
	If sSTATE = "British Columbia" Then sSTATE = "BC"
	If sSTATE = "Manitoba" Then sSTATE = "MB"
	If sSTATE = "New Brunswick" Then sSTATE = "NB"
	If sSTATE = "New South Wales" Then sSTATE = "NS"
	If sSTATE = "Newfoundland" Then sSTATE = "NL"
	If sSTATE = "Northwest Territories" Then sSTATE = "NT"
	If sSTATE = "Nova Scotia" Then sSTATE = "NS"
	If sSTATE = "Nunavut" Then sSTATE = "NU"
	If sSTATE = "Ontario" Then sSTATE = "ON"
	If sSTATE = "Prince Edward Island" Then sSTATE = "PE"
	If sSTATE = "Quebec" Then sSTATE = "QC"
	If sSTATE = "Queensland" Then sSTATE = "QLD"
	If sSTATE = "Saskatchewan" Then sSTATE = "SK"
	If sSTATE = "South Australia" Then sSTATE = "SA"
	If sSTATE = "Tasmania" Then sSTATE = "TAS"
	If sSTATE = "Victoria" Then sSTATE = "VIC"
	If sSTATE = "Western Australia" Then sSTATE = "W"
	If sSTATE = "Yukon Territory" Then sSTATE = "YT"
	If sSTATE = "APO AE" Then sSTATE = "AE"
	If sSTATE = "Northern Territory" Then sSTATE = "NT"
	If sSTATE = "Australian Capital Territory" Then sSTATE = "AC"

	iserror = 0
	
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
	
	conDB.BeginTrans
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
'is this a new attendee or is a duplicate address new attendee then insert
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if Session("ATTENDEE_LIST") = "" Or Session("DUPLICATE") <> "" Then
		strSQL = "INSERT INTO Registration (" & _
					"first_name, last_name, address, city, state, zip, country, email, cell_phone, emergency_phone, comments, grade_level_type, position_type, " & _
					"amount_paid, amount_due, preregistered, share_info, hotel, school_company, regtimestamp) " & _
				"VALUES (" 
				
		If sFIRST_NAME = "" Then 
		  strSQL = strSQL & "NULL, "
		Else
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sFIRST_NAME)) & "', "
		End If
		
		If sLAST_NAME = "" Then 
		  strSQL = strSQL & "NULL, "
		Else   
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sLAST_NAME)) & "', "
		End If
		
		If sADDRESS = "" Then 
		  strSQL = strSQL & "NULL, "
		Else     
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sADDRESS)) & "', "
		End If
		
		If sCITY = "" Then 
		  strSQL = strSQL & "NULL, "
		Else    
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sCITY)) & "', "
		End If
		
		If sSTATE = "" Then 
		  strSQL = strSQL & "NULL, "
		Else    
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sSTATE)) & "', "
		End If
		
		If sZIP = "" Then 
		  strSQL = strSQL & "NULL, "
		Else    
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sZIP)) & "', " 
		End If
		
		If sCOUNTRY = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sCOUNTRY)) & "', "
		End If
		
		If sEMAIL = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sEMAIL)) & "', "
		End If
		
		If sCELL_PHONE = "" Then
			strSQL = strSQL & "NULL, "
		Else
			strSQL = strSQL & "'" & Trim(ManageQuotes(sCELL_PHONE)) & "', "
		End If
		
		If sEMERGENCY_PHONE = "" Then
			strSQL = strSQL & "NULL, "
		Else
			strSQL = strSQL & "'" & Trim(ManageQuotes(sEMERGENCY_PHONE)) & "', "
		End If				
		
		If sCOMMENTS = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sCOMMENTS)) & "', "
		End If 
		
		If cGRADE_LEVELS = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & cGRADE_LEVELS & "', "
		End If
		
		If cPOSITION = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & cPOSITION & "', "
		End If
		
		If iRegPayment < 0 Then 
		  strSQL = strSQL & "0, "
		Else 
		  strSQL = strSQL & iRegPayment  & ","
		End If
		
		If sAMOUNT_DUE = "" Then 
		  strSQL = strSQL & "0, 'N',"
		Else 
'the N is for preregistered      
		  strSQL = strSQL & Trim(sAMOUNT_DUE) & ", 'N',"
		End If
		
		If cSHARE_INFO = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & ManageQuotes(cSHARE_INFO) & "', "
		End If
		
		If sHOTEL = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sHOTEL)) & "', "
		End If
		
		If sSCHOOL_COMPANY = "" Then 
		  strSQL = strSQL & "NULL, "
		Else       
		  strSQL = strSQL & "'" & Trim(ManageQuotes(sSCHOOL_COMPANY)) & "', "
		End If
			    		  
		strSQL = strSQL & "GETDATE())"
		
		conDB.Execute strSQL, lngRecs, adCmdtext
	
		If lngRecs <> 1 Then
			iserror = iserror + 1
		End if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
'get the new reg_id and insert into badge, badge_activity, position, payment, program
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		strSQL = "SELECT registration_id FROM registration " & _
				 "WHERE first_name =  '" & Trim(ManageQuotes(sFIRST_NAME)) & "' " & _
				 "AND last_name = '" & Trim(ManageQuotes(sLAST_NAME)) & "' " & _
				 "and CITY = '" & Trim(ManageQuotes(sCITY)) & "' "
				 
		set rsRegID = conDB.Execute(strSQL, lngRecs, adCmdtext)	

		newRegID = rsRegID("registration_id")
		
		Session("REGISTRATION_ID") = newRegID
		
		Set rsRegID = Nothing	
		
		strSQL = "INSERT INTO Badge (" & _
					"badge_type, registration_id, badge_last_name, badge_first_name, complimentary, badge_company, badge_city, badge_state, badge_country)" & _
				 "VALUES (" 
				 
		strSQL = strSQL & "'" & cBADGE_TYPE & "', "		
		
		strSQL = strSQL & newRegID & ", '"
		
		strSQL = strSQL & Trim(ManageQuotes(sLAST_NAME)) & "', '"
		
		strSQL = strSQL & Trim(ManageQuotes(sFIRST_NAME)) & "', "
		
		If cBADGE_COMP = "" Then 
		  strSQL = strSQL & "'N', '"
		Else       
		  strSQL = strSQL & "'" & ManageQuotes(cBADGE_COMP) & "', '"
		End If
		
		strSQL = strSQL & Trim(ManageQuotes(sSCHOOL_COMPANY)) & "', '"
		
		strSQL = strSQL & Trim(ManageQuotes(sCITY)) & "', '"
		
		strSQL = strSQL & Trim(ManageQuotes(sSTATE)) & "', '"
		
		strSQL = strSQL & Trim(ManageQuotes(sCOUNTRY)) & "')"
		
		conDB.Execute strSQL, lngRecs, adCmdtext
		
		If lngRecs <> 1 Then
			iserror = iserror + 1
		End if
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
'if extra programs were added, insert here
''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
		If sADDL_PROG_COUNT <> "" or aADDL_PROG_COUNT <> "0" Then 
			strSQL = "INSERT INTO program (registration_id, num_programs, payment_type, acttimestamp) VALUES (" 
			
			strSQL = strSQL & newRegID & ", " & CInt(sADDL_PROG_COUNT) & ", "
			
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
		
		If cPRINT_BADGE <> "" Then
'get the new badge_id and insert into log_badge_activity, log_badge_payment_activity
			Session("PRINT_ATTENDEE") = "Y"
			
			strSQL = "SELECT badge_id FROM badge " & _
					"WHERE registration_id = " & newRegID 
					
			Set rsBadgeID = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
			newBadgeID = rsBadgeID("badge_id")
			
			Set rsBadgeID = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' includes who picked them up
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			strSQL = "INSERT INTO log_badge_activity (" & _
						"acttimestamp, badge_id, registration_id) " & _
					"VALUES ("
					
			strSQL = strSQL & "GETDATE(), " & newBadgeID & ", " & newRegID & ")"

			conDB.Execute strSQL, lngRecs, adCmdtext 
		
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if			
		End If
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' only insert into the log_badge_payment_activity if there
' is a payment (not just a program payment)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		If iRegPayment > 0 Then	
			strSQL = "INSERT INTO log_badge_payment_activity (" & _
						"acttimestamp, registration_id, amount_paid, payment_type, complimentary, tot_amount_due, initials) " & _
					 "VALUES ("
					 
			strSQL = strSQL & "GETDATE(), " & newRegID & ", "
			
			If iRegPayment < 0 Then
			  strSQL = strSQL & "0, "
			Else       
			  strSQL = strSQL & iRegPayment & ", "
			End If
			
			If cPAYMENT_TYPE = "" Then 
			  strSQL = strSQL & "NULL, "
			Else       
			  strSQL = strSQL & "'" & cPAYMENT_TYPE & "', "
			End If
			
			If cBADGE_COMP = "" Then 
			  strSQL = strSQL & "'N', "
			Else       
			  strSQL = strSQL & "'" & ManageQuotes(cBADGE_COMP) & "', "
			End If
			
			If sAMOUNT_DUE = "" Then 
			  strSQL = strSQL & "0, "
			Else  
			  strSQL = strSQL & Trim(sAMOUNT_DUE) & ", "
			End If
			
			If sBADGE_INITIALS = "" Then 
			  strSQL = strSQL & "NULL)"
			Else       
			  strSQL = strSQL & "'" & ManageQuotes(sBADGE_INITIALS) & "')"
			End If		
			
			conDB.Execute strSQL, lngRecs, adCmdtext
					
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End If			
		End If

		Session("PREVIOUS_ATTENDEE") = Trim(ManageQuotes(sLAST_NAME)) & ", " & Trim(ManageQuotes(sFIRST_NAME)) & ", " & Trim(ManageQuotes(sCITY)) & ", " & Trim(ManageQuotes(sSTATE))
		Session("ATTENDEE_LIST") = ""
		
		Session("DUPLICATE") = ""
	Else
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update
' parse for first, last name, city, and state and get the 
' reg_id we want to update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		arAttendeeField = split(Session("ATTENDEE_LIST"), " | ")
		
		sessionLast = arAttendeeField(0)
		
		sessionFirst = arAttendeeField(1)
		
		If arAttendeeField(2) = "|" then
			sessionCity = ""
		ElseIf arAttendeeField(2) = "" or arAttendeeField(2) = null Then
			sessionCity = ""
		ElseIf Right(arAttendeeField(2),1) = "|" Then				
			sessionCity = Mid(arAttendeeField(2),1,Len(arAttendeeField(2)) - 1)	
		Else
			sessionCity = arAttendeeField(2)
		End If

		sessionCity = Replace(sessionCity,"'","''")
		
		If InStr(1,arAttendeeField(2),"|") > 0 then
			sessionState = ""
		Elseif arAttendeeField(3) = "" or arAttendeeField(3) = null then
			sessionState = ""
		Else
			sessionState = arAttendeeField(3)		
		End If
		
		If InStr(1, sessionLast, "^") Then
			sPre = "^"
			
			sessionLast = Trim(Mid(sessionLast, 1, len(sessionlast) -1))
		End If
		
		If sPre = "^" Then
			strSQL = "SELECT registration_id FROM badge " & _
					 "WHERE badge_first_name = '" & ManageQuotes(Trim(sessionFirst)) & "' " & _
					 "AND badge_last_name = '" & ManageQuotes(Trim(sessionLast)) & "' " & _ 
					 "AND badge_city = '" & sessionCity & "' " & _ 
					 "AND badge_state = '" & sessionState & "' " 

			Set rsRegID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			iRegID = rsRegID("registration_id")
			
			Session("REGISTRATION_ID") = iRegID
			
			Set rsRegID = Nothing
			
'update the name/city/state in badge
			strSQL = "UPDATE badge SET badge_first_name = "
			 
			If sFIRST_NAME = "" Then 
			   strSQL = strSQL & "NULL, "
			Else
			   strSQL = strSQL & "'" & Trim(ManageQuotes(sFIRST_NAME)) & "', "
			End If
			
			strSQL = strSQL & "badge_last_name = "
			
			If sLAST_NAME = "" Then 
			   strSQL = strSQL & "NULL, "
			Else   
			   strSQL = strSQL & "'" & Trim(ManageQuotes(sLAST_NAME)) & "', "
			End If	
			
			strSQL = strSQL & "badge_city = "

			If sCITY = "" or sCITY = null Then
			   strSQL = strSQL & "NULL, "
			Else
				strSQL = strSQL & "'" & Trim(ManageQuotes(sCITY)) & "', "
			End If	
			
			strSQL = strSQL & "badge_state = "
			
			If sSTATE = "" or len(sSTATE) = 1 or sSTATE = null Then
				strSQL = strSQL & "NULL, "
			Else
				strSQL = strSQL & "'" & Trim(ManageQuotes(sSTATE)) & "', "
			End If			

			strSQL = strSQL & "WHERE registration_id = " & iRegID
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
' then update the rest in registration
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			strSQL = "UPDATE registration SET address = "
			
			If sADDRESS = "" Then 
				strSQL = strSQL & "NULL, "
			Else     
				strSQL = strSQL & "'" & Trim(ManageQuotes(sADDRESS)) & "', "
			End If	
			
			strSQL = strSQL & "zip = "
			
			If sZIP = "" Then 
			    strSQL = strSQL & "NULL, "
			Else    
			    strSQL = strSQL & "'" & Trim(ManageQuotes(sZIP)) & "', " 
			End If	
			
			strSQL = strSQL & "country = "
			
			If sCOUNTRY = "" Then 
			    strSQL = strSQL & "NULL, "
			Else       
			    strSQL = strSQL & "'" & Trim(ManageQuotes(sCOUNTRY)) & "', "
			End If
			
			strSQL = strSQL & "email = "
			
			If sEMAIL = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sEMAIL)) & "', "
			End If
			
			strSQL = strSQL & "cell_phone = "
			
			If sCELL_PHONE = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & Trim(ManageQuotes(sCELL_PHONE)) & "', "
			End If
			
			strSQL = strSQL & "emergency_phone = "
			
			If sEMERGENCY_PHONE = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & Trim(ManageQuotes(sEMERGENCY_PHONE)) & "', "
			End If
			
			strSQL = strSQL & "comments = "
			
			If sCOMMENTS = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sCOMMENTS)) & "', "
			End If
			
			strSQL = strSQL & "grade_level_type = "
			
			If cGRADE_LEVELS = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & cGRADE_LEVELS & "', "
			End If
			
			strSQL = strSQL & "position_type = "
			
			If cPOSITION = "" Then 
				strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & cPOSITION & "', "
			End If
			
			strSQL = strSQL & "amount_paid = "
			
			If iRegPayment < 0 Then 
				strSQL = strSQL & "0, "
			Else       
				strSQL = strSQL & iRegPayment & ", "
			End If
			
			strSQL = strSQL & "amount_due = "
			
			If sAMOUNT_DUE = "" Then 
				strSQL = strSQL & "0, "
			Else       
				strSQL = strSQL & Trim(sAMOUNT_DUE) & ", "
			End If
			
			strSQL = strSQL & "share_info = "
			
			If cSHARE_INFO = "" Then 
				strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'No', "
			End If
			
			strSQL = strSQL & "hotel = "
			
			If sHOTEL = "" Then 
				strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & Trim(ManageQuotes(sHOTEL)) & "', "
			End If
			
			strSQL = strSQL & "school_company = "
			
			If sSCHOOL_COMPANY = "" Then 
				strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & Trim(ManageQuotes(sSCHOOL_COMPANY)) & "', "
			End If	
					
			strSQL = strSQL & "regtimestamp = GETDATE() WHERE registration_id = " & iRegID
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		Else
			If InStr(1, sessionLast,"**") Then
				sessionLast = Mid(ManageQuotes(Trim(sessionLast)), 4)
			Else
				sessionLast = ManageQuotes(Trim(sessionLast))
			End If			
			
			strSQL = "SELECT registration_id FROM registration " & _
					 "WHERE first_name = '" & ManageQuotes(Trim(sessionFirst)) & "' " & _
					 "AND last_name = '" & sessionLast & "' "
					 
			If sessionCity <> "" then
				strSQL = strSQL & "AND city = '" & sessionCity & "' "
			End If
			
			Set rsRegID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			If rsRegID.EOF then
				strSQL = "SELECT registration_id FROM Badge " & _
					 "WHERE badge_first_name = '" & ManageQuotes(Trim(sessionFirst)) & "' " & _
					 "AND badge_last_name = '" & sessionLast & "' "
					 
				If sessionCity <> "" then
					strSQL = strSQL & "AND badge_city = '" & sessionCity & "' "
				End If

				Set rsRegID = conDB.Execute(strSQL, lngRecs, adCmdtext)			
			End IF
			
			iRegID = rsRegID("registration_id")
			
			Session("REGISTRATION_ID") = iRegID
			
			Set rsRegID = Nothing
			
			strSQL = "UPDATE registration SET first_name = "
			
			If sFIRST_NAME = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sFIRST_NAME)) & "', "
			End If
			
			strSQL = strSQL & "last_name = "
			
			If sLAST_NAME = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else   
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sLAST_NAME)) & "', "
			End If	
			
			strSQL = strSQL & "address = "
			
			If sADDRESS = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else     
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sADDRESS)) & "', "
			End If	
			
			strSQL = strSQL & "city = "
			
			If sCITY = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else    
			   	strSQL = strSQL & "'" & sCity & "', "
			End If	
			
			strSQL = strSQL & "state = "
			
			If sSTATE = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else
			   	strSQL = strSQL & "'" & sState & "', "
			End If
			
			strSQL = strSQL & "zip = "
			
			If sZIP = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else    
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sZIP)) & "', " 
			End If	
			
			strSQL = strSQL & "country = "
			
			If sCOUNTRY = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sCOUNTRY)) & "', "
			End If
			
			strSQL = strSQL & "email = "
			
			If sEMAIL = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   strSQL = strSQL & "'" & Trim(ManageQuotes(sEMAIL)) & "', "
			End If
			
			strSQL = strSQL & "cell_phone = "
			
			If sCELL_PHONE = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sCELL_PHONE)) & "', "
			End If
			
			strSQL = strSQL & "emergency_phone = "
			
			If sEMERGENCY_PHONE = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & Trim(ManageQuotes(sEMERGENCY_PHONE)) & "', "
			End If
			
			strSQL = strSQL & "comments = "
			
			If sCOMMENTS = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
				strSQL = strSQL & "'" & Trim(ManageQuotes(sCOMMENTS)) & "', "
			End If
			
			strSQL = strSQL & "grade_level_type = "
			
			If cGRADE_LEVELS = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & cGRADE_LEVELS & "', "
			End If
			
			strSQL = strSQL & "position_type = "
			
			If cPOSITION = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'" & cPOSITION & "', "
			End If
			
			strSQL = strSQL & "amount_due = "
			
			If sAMOUNT_DUE = "" Then 
			   	strSQL = strSQL & "0, "
			Else       
			   	strSQL = strSQL & Trim(sAMOUNT_DUE) & ", "
			End If
			
			strSQL = strSQL & "share_info = "
			
			If cSHARE_INFO = "" Then 
			   	strSQL = strSQL & "NULL, "
			Else       
			   	strSQL = strSQL & "'No', "
			End If
			
			strSQL = strSQL & "hotel = "
			
			If sHOTEL = "" Then 
			  	strSQL = strSQL & "NULL, "
			Else       
			  	strSQL = strSQL & "'" & Trim(ManageQuotes(sHOTEL)) & "', "
			End If
			
			strSQL = strSQL & "school_company = "
			
			If sSCHOOL_COMPANY = "" Then 
			  	strSQL = strSQL & "NULL, "
			Else       
			  	strSQL = strSQL & "'" & Trim(ManageQuotes(sSCHOOL_COMPANY)) & "', "
			End If		
				
			strSQL = strSQL & "regtimestamp = GETDATE() WHERE registration_id = " & iRegID
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if	
		End If
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' get the badge_id and update badge, log_badge_activity, log_badge_payment_activity and program
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		strSQL = "SELECT badge_id, badge_last_name, badge_first_name FROM badge " & _
				 "WHERE registration_id = " & iRegID 

		Set rsBadgeID = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		iBadgeID = rsBadgeID("badge_id")
		
		Set rsBadgeID = Nothing
		
		strSQL = "UPDATE Badge SET badge_type = " 
		
		If cBADGE_TYPE = "" Then 
		   	strSQL = strSQL & "NULL, "
		Else       
		   	strSQL = strSQL & "'" & Trim(ManageQuotes(cBADGE_TYPE)) & "', "
		End If
		
		strSQL = strSQL & "complimentary = "
		
		If cBADGE_COMP = "" Then 
		  	strSQL = strSQL & "'N', "
		Else       
		  	strSQL = strSQL & "'" & ManageQuotes(cBADGE_COMP) & "', "
		End If
		
		strSQL = strSQL & "badge_first_name = "
		
		If sFIRST_NAME = "" Then 
			strSQL = strSQL & "NULL, "
		Else
			strSQL = strSQL & "'" & Trim(ManageQuotes(sFIRST_NAME)) & "', "
		End If
		
		strSQL = strSQL & "badge_last_name = "
		
		If sLAST_NAME = "" Then 
			strSQL = strSQL & "NULL, "
		Else   
			strSQL = strSQL & "'" & Trim(ManageQuotes(sLAST_NAME)) & "', "
		End If	
		
		strSQL = strSQL & "badge_city = "
		
		If sCITY = "" Then 
			strSQL = strSQL & "NULL, "
		Else    
			strSQL = strSQL & "'" & sCity & "', "
		End If	
		
		strSQL = strSQL & "badge_state = "
		
		If sSTATE = "" Then 
			strSQL = strSQL & "NULL, "
		Else
			strSQL = strSQL & "'" & sState & "', "
		End If
		
		strSQL = strSQL & "badge_country = "
		
		If sCOUNTRY = "" Then 
			strSQL = strSQL & "NULL, "
		Else       
			strSQL = strSQL & "'" & Trim(ManageQuotes(sCOUNTRY)) & "', "
		End If
		
		
		strSQL = strSQL & "badge_company = "
		
		If sCOUNTRY = "" Then 
			strSQL = strSQL & "NULL "
		Else       
			strSQL = strSQL & "'" & Trim(ManageQuotes(sCOUNTRY)) & "' "
		End If
		
		strSQL = strSQL & "WHERE badge_id = " & iBadgeID
		
		
		conDB.Execute strSQL, lngRecs, adCmdtext 
		 
		If lngRecs <> 1 Then
			iserror = iserror + 1
		End if
		

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if there was nothing in table program, insert it.  if there is an entry, then update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		if sADDL_PROG_COUNT <> "0" then		
			strSQL = "SELECT count(*) Cnt FROM program WHERE registration_id = " & iRegID 
			
			Set rsProgID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			iCnt = rsProgID("Cnt")
			
			Set rsProgID = Nothing
		
			If iCnt = 0 Then
				strSQL = "INSERT INTO program (registration_id, num_programs, payment_type, acttimestamp) VALUES (" 
				
				strSQL = strSQL & iRegID & ", " & CInt(sADDL_PROG_COUNT) & ", "
				
				If cPAYMENT_TYPE = "" Then 
				  strSQL = strSQL & "NULL, GETDATE())"
				Else       
				  strSQL = strSQL & "'" & cPAYMENT_TYPE & "', GETDATE())"
				End If
				
				conDB.Execute strSQL, lngRecs, adCmdtext 
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
			Else 
				strSQL = "UPDATE Program SET num_programs = " & sADDL_PROG_COUNT
				
				strSQL = strSQL & ", acttimestamp = GETDATE() WHERE registration_id = " & iRegID
				conDB.Execute strSQL, lngRecs, adCmdtext 
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
			End If
		end if
	
		If cPRINT_BADGE <> "" Then
'includes who picked them up
			strSQL = "INSERT INTO log_badge_activity (" & _
						"acttimestamp, badge_id, registration_id, pickup_name) " & _
					"VALUES ("
					
			strSQL = strSQL & "GETDATE(), " & iBadgeID & ", " & iRegID & ", '" & ManageQuotes(sBADGE_PICKUP_NAME) & "')"
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		End If 
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
' if the entries into log_badge_payment_activity total
' the amount due, then don't insert
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		strSQL = "select sum(amount_paid) amount_paid, tot_amount_due " & _
		         "from log_badge_payment_activity " & _
		         "where registration_id = " & iRegID & _ 
		         "group by tot_amount_due"
				 
		Set rsBadgeID = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		If Not rsBadgeID.EOF Then
			sumAmountPaid = rsBadgeID("amount_paid")
		Else
			sumAmountPaid = 0
		End If
		
		Set rsBadgeID = Nothing
		
		If CInt(sumAmountPaid) < CInt(sAMOUNT_DUE) Then		
			strSQL = "INSERT INTO log_badge_payment_activity (" & _
						"acttimestamp, registration_id, amount_paid, payment_type, complimentary, tot_amount_due, initials) " & _
					 "VALUES ("
					 
			strSQL = strSQL & "GETDATE(), " & iRegID & ", "
			
			If iRegPayment < 0 Then
			  	strSQL = strSQL & "0, "
			Else       
			  	strSQL = strSQL & iRegPayment & ", "
			End If
			
			If cPAYMENT_TYPE = "" Then 
			  	strSQL = strSQL & "NULL, "
			Else       
			  	strSQL = strSQL & "'" & cPAYMENT_TYPE & "', "
			End If
			
			If cBADGE_COMP = "" Then 
			  	strSQL = strSQL & "'N', "
			Else       
			  	strSQL = strSQL & "'" & ManageQuotes(cBADGE_COMP) & "', "
			End If
			
			If sAMOUNT_DUE = "" Then 
			 	strSQL = strSQL & "0, "
			Else  
			  	strSQL = strSQL & Trim(sAMOUNT_DUE) & ", "
			End If
			
			If sBADGE_INITIALS = "" Then 
			  	strSQL = strSQL & "NULL)"
			Else       
			  	strSQL = strSQL & "'" & ManageQuotes(sBADGE_INITIALS) & "')"
			End If		
	
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End If
			
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' now update registration with the new total paid
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			strSQL = "UPDATE REGISTRATION SET amount_paid = " & (CInt(sumAmountPaid) + iRegPayment) & " " & _
			         "WHERE registration_id = " & iRegID
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if			
		End If
				
		Session("PREVIOUS_ATTENDEE") = Session("ATTENDEE_LIST")
		
		Session("ATTENDEE_LIST") = ""		
	End If
	
	If iserror = 0 Then
		conDB.CommitTrans
	Else
		conDB.RollbackTrans
	End If
	
	conDB.Close
	
	Set conDB = Nothing

	
	s75 = "DLG"
	s25 = "S"
	sZero = "CF"

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon

	strSQL = "select r.registration_id, r.first_name, r.last_name, r.city, r.state, r.country, r.school_company, " & _
			 "b.badge_type, b.badge_first_name, b.badge_last_name, b.badge_city, b.badge_state, b.complimentary, preregistered " & _
			 "from registration r, badge b " & _
			 "where r.registration_id = b.registration_id " & _
			 "and r.registration_id = " & Session("REGISTRATION_ID")

	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
	
	If Not rs.EOF Then	
		If rs("badge_first_name") <> vbnullstring and rs("badge_last_name") <> vbnullstring Then
			sBadgeName = Mid(rs("badge_first_name") & "  " & rs("badge_last_name"), 1, 27)
		Else
			sBadgeName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 27)
		End If
		
		sSchoolCompany = rs("school_company")
		
'split into 2 rows if large
        If Len(sSchoolCompany) > 30 Then
'find the closest space before character 26, get it's number
            s = Mid(sSchoolCompany, 1, 31)  's = United States Air Force Ba
            
			For i = 31 To 1 Step -1
                t = StrComp(Mid(s, i, 1), Chr(32))
                
				If t = 0 Then
                    sSC1 = Mid(sSchoolCompany, 1, i)
                    
					sSC2 = Mid(sSchoolCompany, i + 1, Len(sSchoolCompany))
                    
					Exit For
                End If
            Next
'place all text before it in sSC1, and all after it in sSC2
        Else
            sSC1 = sSchoolCompany
			
			sSC2 = "&nbsp"
        End If		

		If rs("badge_city") <> vbnullstring and rs("badge_state") <> vbnullstring Then
			sBadgeFrom = Mid(rs("badge_city") & "  " & rs("badge_state"), 1, 30)
		Else
			sBadgeFrom = Mid(rs("city") & "  " & rs("state"), 1, 30)
		End If
	
		sName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 27)
		
		sFrom = Mid(rs("city") & ", " & rs("state"), 1, 30)
		
		If rs("country") = "United States" Then
			sCountry = ""
		Else
			sCountry = Mid(rs("country"), 1, 27)
		End If
		
		If rs("badge_type") = "D" Then
			btype = "DIRECTOR/ADULT"
		ElseIf rs("badge_type") = "L" Then
			btype = "COMMERCIAL"
		ElseIf rs("badge_type") = "G" Then
			btype = "GUEST"
		ElseIf rs("badge_type") = "S" Then
			btype = "COLLEGE STUDENT"
		ElseIf rs("badge_type") = "C" Then
			btype = "CLINICIAN"
		ElseIf rs("badge_type") = "F" Then
			btype = "STAFF"
		End If
		
		If rs("preregistered") = "Y" Then
			sPayType = "Preregistration"
		Else
			sPayType = "Registration"
		End If
		
		If rs("complimentary") = "Y" Or instr(sZero, rs("badge_type")) Then
			sAmount = " 0.00"
		ElseIf instr(s75, rs("badge_type")) and sPayType = "Preregistration" Then
			sAmount = "110.00"
		ElseIf instr(s75, rs("badge_type")) and sPayType = "Registration" Then
			sAmount = "135.00"
		ElseIf instr(s25, rs("badge_type")) and sPayType = "Registration" Then
			sAmount = "75.00"
		ElseIf instr(s25, rs("badge_type")) and sPayType = "Preregistration" Then
			sAmount = "50.00"
		End If
		
		sBadgeString = "^XA^A0N,100,100^FO0,120^FB810,3,,C^FD" & sBadgeName 


		if len(sBadgeName) > 17 then			
		sBadgeString = sBadgeString & "^FS^A0N,50,50^FO0,310^FB810,3,,C^FD" & sSC1 & "^FS" & _
"^A0N,70,70^FO0,370^FB816,1,,C^FD" & sBadgeFrom & "^FS" & _
"^A0N,70,70^FO0,440^FB816,1,,C^FD" & sCountry & "^FS" & _
"^A0N,50,50^FO0,480^FB816,1,,C^FD" & btype & "^FS" & _
"^XZ"
else
sBadgeString = sBadgeString & "^FS^A0N,50,50^FO0,210^FB810,3,,C^FD" & sSC1 & "^FS" & _
"^A0N,70,70^FO0,270^FB816,1,,C^FD" & sBadgeFrom & "^FS" & _
"^A0N,70,70^FO0,340^FB816,1,,C^FD" & sCountry & "^FS" & _
"^A0N,50,50^FO0,480^FB816,1,,C^FD" & btype & "^FS" & _
"^XZ"
End if

		sReceiptString= "^XA" & _
"^A0N,40,40^FO40,100^FD" & sName & "^FS" & _
"^A0N,40,40^FO680,100^FD" & rs("registration_id") & "^FS" & _
"^A0N,40,40^FO40,150^FDTOTAL PAID:^FS" & _
"^A0N,40,40^FO680,150^FD" & sAmount & "^FS" & _
"^A0N,40,40^FO40,240^FDREGISTRATION PAYMENT DETAILS:^FS" & _
"^A0N,40,40^FO40,290^FDPAY TYPE^FS" & _
"^A0N,40,40^FO350,290^FDDATE^FS" & _
"^A0N,40,40^FO650,290^FDAMOUNT^FS" & _
"^A0N,40,40^FO40,360^FD" & sPayType & "^FS" & _
"^A0N,40,40^FO320,360^FD" & Date & "^FS" & _
"^A0N,40,40^FO650,360^FD" & sAmount & "^FS" & _
"^XZ"
	End If
	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing
	
	If cPRINT_BADGE <> "" Then
				printobject = printobject & "Set objFS = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine
			printobject = printobject & "Set objPrint = objFS.CreateTextFile(""LPT1:"", True)" & vbNewLine
			printobject = printobject & "objPrint.Write(""" & sBadgeString & """)" & vbNewLine
			printobject = printobject & "objPrint.Close" & vbNewLine
			
						printobject = printobject & "Set objFS = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine
			printobject = printobject & "Set objPrint = objFS.CreateTextFile(""LPT1:"", True)" & vbNewLine
			printobject = printobject & "objPrint.Write(""" & sReceiptString & """)" & vbNewLine
			printobject = printobject & "objPrint.Close" & vbNewLine
	End if
%>
<html>
	<body>
		<script type="text/vbscript">
			<%= printobject%>

			window.location = "registration.asp"
		</script>
	</body>
</html>