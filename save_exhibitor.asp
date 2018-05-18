<% @Language=VBScript %>
<% Response.Buffer="true"%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sBadgeString
	dim sReceiptString
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
	dim conDB
	dim strSQL
	dim lngRecs
	dim rsBadgeID
	dim rsNumBadges
	dim iNumBadges
	dim iBadgeID
	dim rsGuestID
	dim iGuestID
	dim iBoothAmtPd
	dim iBoothAmtOut
	dim iTotBoothAmtPd
	dim rsAddlComp
	dim iAddlComp
	dim iNewNonComp
	dim iBadgeAmtPd
	dim iBadgeAmtOut
	dim iTotBadgeAmtPd
	dim iserror
	dim idx
	dim arBadgeNums
	dim iBadgeNum
	dim sPrintBadgeString
	dim sFirstName
	dim sLastName
	dim sCOMPANY_NAME
	dim sCONTACT_NAME	
	dim sEMAIL	
	dim sZIP	
	dim sCITY		
	dim sSTATE		
	dim sCOUNTRY		
	dim sADDRESS		
	dim sPHONE	
	dim sBOOTH_AMOUNT_DUE	
	dim sBOOTH_PAYMENT	
	dim cBOOTH_PAYMENT_TYPE
	dim sBOOTH_INITIALS
	dim sNUM_COMP_BADGES	
	dim sNUM_COMP		
	dim sADDL_COMP	
	dim sORIG_NUM_COMP	
	dim sNUM_NON_COMP		
	dim sBADGE_AMOUNT_DUE	
	dim sNUM_REMAINING_BADGES	
	dim sBADGE_PAYMENT	
	dim cBADGE_PAYMENT_TYPE	
	dim sBADGE_INITIALS	
	dim sBADGE_PICKUP_NAME	
	dim cPRINTBADGE		
	dim cGOTTICKETS		
	dim cINCPAYMENT
	dim cINCPROGRAMPAYMENT
	dim cProgramCount
	dim sCOMMENTS
	Dim printobject
	Dim printobject2

	Session("Test") = "Test2"
	sCOMPANY_NAME		= Trim(Request("COMPANY_NAME"))
	sCONTACT_NAME		= Trim(Request("CONTACT_NAME"))
	sEMAIL				= Trim(Request("EMAIL"))
	sZIP				= Trim(Request("ZIP"))
	sCITY				= Trim(Request("CITY"))
	sSTATE				= Trim(Request("STATES"))
	sCOUNTRY			= Trim(Request("COUNTRY"))
	sADDRESS			= Trim(Request("ADDRESS"))
	sPHONE				= Trim(Request("PHONE"))
	sBOOTH_AMOUNT_DUE	= Trim(Request("BOOTH_AMOUNT_DUE"))
	sBOOTH_PAYMENT		= Trim(Request("BOOTH_PAYMENT"))
	cBOOTH_PAYMENT_TYPE	= Trim(Request("BOOTH_PAYMENT_TYPE"))
	sBOOTH_INITIALS		= Trim(Request("BOOTH_INITIALS"))
	iTotBoothAmtPd		= Trim(Request("TOTAL_BOOTH_AMOUNT_PAID"))
	sNUM_COMP_BADGES	= Trim(Request("NUM_COMP_BADGES"))
	sNUM_COMP			= Trim(Request("NUM_COMP"))
	sADDL_COMP			= Trim(Request("ADDL_COMP"))
	sORIG_NUM_COMP		= Trim(Request("ORIG_NUM_COMP"))
	sNUM_NON_COMP		= Trim(Request("NUM_NON_COMP"))
	sBADGE_AMOUNT_DUE	= Trim(Request("BADGE_AMOUNT_DUE"))
	sNUM_REMAINING_BADGES	= Trim(Request("NUM_REMAINING_BADGES"))
	sBADGE_PAYMENT		= Trim(Request("BADGE_PAYMENT"))
	cBADGE_PAYMENT_TYPE	= Trim(Request("BADGE_PAYMENT_TYPE"))
	sBADGE_INITIALS		= Trim(Request("BADGE_INITIALS"))
	sBADGE_PICKUP_NAME	= Trim(Request("BADGE_PICKUP_NAME"))
	iTotBadgeAmtPd		= Trim(Request("TOTAL_BADGE_AMOUNT_PAID"))
	cPRINTBADGE			= Trim(Request("PRINTBADGE"))
	cGOTTICKETS			= Trim(Request("GOTTICKETS"))
	cINCPAYMENT			= Trim(Request("INCPAYMENT"))
	cINCPROGRAMPAYMENT	= Trim(Request("INCPROGRAMPAYMENT"))
	sCOMMENTS			= Trim(Request("COMMENTS"))
	printobject = ""

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

	If Len(Request.Form("SORT")) Then
		If Request.Form("SORT_BY") <> "" Then
			Session("SORT_ORDER") = Request.Form("SORT_BY")
		Else
			Session("SORT_ORDER") = "badge_num"
		End if
	Else
	
		Session("PREVIOUS_EXHIBITOR") = Session("EXHIBITOR_ID")
		
		If iTotBoothAmtPd = "" Then
			iTotBoothAmtPd = sBOOTH_AMOUNT_DUE
		End If
		
		If iTotBadgeAmtPd = "" Then
			iTotBadgeAmtPd = sBADGE_AMOUNT_DUE
		End If
		
		iserror = 0
	
		Set conDB = Server.CreateObject("ADODB.Connection")
		
		conDB.Open strCon
		
		conDB.BeginTrans
	
		strSQL = "UPDATE exhibitor SET company_name = " 
		
		If sCOMPANY_NAME = "" Then 
		   strSQL = strSQL & "NULL, "
		Else
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sCOMPANY_NAME)) & "', "
		End If
		
		strSQL = strSQL & "contact_name = "
		
		If sCONTACT_NAME = "" Then 
		   strSQL = strSQL & "NULL, "
		Else     
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sCONTACT_NAME)) & "', "
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
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sCITY)) & "', "
		End If	
		
		strSQL = strSQL & "state = "
		
		If sSTATE = "" Then 
		   strSQL = strSQL & "NULL, "
		Else    
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sSTATE)) & "', "
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
		
		strSQL = strSQL & "phone = "
		
		If sPHONE = "" Then 
		   strSQL = strSQL & "NULL, "
		Else       
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sPHONE)) & "', "
		End If
		
		strSQL = strSQL & "booth_amount_due = "
		
		If sBOOTH_AMOUNT_DUE = "" Then 
		   strSQL = strSQL & "0, "
		Else       
		   strSQL = strSQL & Trim(sBOOTH_AMOUNT_DUE) & ", "
		End If
		
		strSQL = strSQL & "badge_amount_due = "
		
		If sBADGE_AMOUNT_DUE = "" Then 
		   strSQL = strSQL & "0, "
		Else       
		   strSQL = strSQL & Trim(sBADGE_AMOUNT_DUE) & ", "
		End If
		
		strSQL = strSQL & "comments = "
		
		If sCOMMENTS = "" Then 
		   strSQL = strSQL & "NULL, "
		Else       
		   strSQL = strSQL & "'" & Trim(ManageQuotes(sCOMMENTS)) & "', "
		End If
		
		strSQL = strSQL & "regtimestamp = GETDATE() WHERE exhibitor_id = " & Session("EXHIBITOR_ID")
		
		conDB.Execute strSQL, lngRecs, adCmdtext 
		
		If lngRecs <> 1 Then
 			iserror = iserror + 1
 		End if
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
'is there a booth payment, if so, insert into the activity table
'but if existing payments total amount due, don't insert
''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

		If sBOOTH_PAYMENT <> "" Then		
			strSQL = "select sum(amount_paid) amount_paid, tot_amount_due " & _
			         "from log_booth_payment_activity " & _
			         "where exhibitor_id = " & Session("EXHIBITOR_ID") & _ 
			         " group by tot_amount_due"
					 
			Set rsBooth = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			If Not rsBooth.EOF Then
				sumAmountPaid = rsBooth("amount_paid")
			Else
				sumAmountPaid = 0
			End If	
			
			Set rsBooth = Nothing
			
			If CInt(sumAmountPaid) < CInt(sBOOTH_AMOUNT_DUE) Then				
				strSQL = "INSERT INTO log_booth_payment_activity (" & _
							"acttimestamp, exhibitor_id, amount_paid, payment_type, tot_amount_due, initials) " & _
						 "VALUES ("
						 
				strSQL = strSQL & "GETDATE(), " & Session("EXHIBITOR_ID") & ", "
				
				strSQL = strSQL & sBOOTH_PAYMENT & ", "
				
				If cBOOTH_PAYMENT_TYPE = "" Then 
				  strSQL = strSQL & "NULL, "
				Else       
				  strSQL = strSQL & "'" & cBOOTH_PAYMENT_TYPE & "', "
				End If
				
				If sBOOTH_AMOUNT_DUE = "" Then 
				  strSQL = strSQL & "0, "
				Else  
				  strSQL = strSQL & Trim(sBOOTH_AMOUNT_DUE) & ", "
				End If
				
				If sBOOTH_INITIALS = "" Then 
				  strSQL = strSQL & "NULL)"
				Else       
				  strSQL = strSQL & "'" & sBOOTH_INITIALS & "')"
				End If
				
				conDB.Execute strSQL, lngRecs, adCmdtext
			
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End If
			End If
		End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'badge payment (old)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		If sBADGE_PAYMENT <> "" Then
			strSQL = "select sum(amount_paid) amount_paid, tot_amount_due " & _
			         "from log_badge_payment_activity " & _
			         "where exhibitor_id = " & Session("EXHIBITOR_ID") & _ 
			         " group by tot_amount_due"
					 
			Set rsBadgeID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			If Not rsBadgeID.EOF Then
				sumAmountPaid = rsBadgeID("amount_paid")
			Else
				sumAmountPaid = 0
			End If
			
			Set rsBadgeID = Nothing
			
			If CInt(sumAmountPaid) < CInt(sBADGE_AMOUNT_DUE) Then
				strSQL = "INSERT INTO log_badge_payment_activity (" & _
							"acttimestamp, exhibitor_id, amount_paid, payment_type, tot_amount_due, initials) " & _
						 "VALUES ("
						 
				strSQL = strSQL & "GETDATE(), " & Session("EXHIBITOR_ID") & ", "
				
				iBadgeAmtPd = sBADGE_PAYMENT
				
				strSQL = strSQL & sBADGE_PAYMENT & ", "
				
				If cBADGE_PAYMENT_TYPE = "" Then 
				  strSQL = strSQL & "NULL, "
				Else       
				  strSQL = strSQL & "'" & cBADGE_PAYMENT_TYPE & "', "
				End If
				
				If sBADGE_AMOUNT_DUE = "" Then 
				  strSQL = strSQL & "0, "
				Else  
				  strSQL = strSQL & Trim(sBADGE_AMOUNT_DUE) & ", "
				End If
				
				If sBADGE_INITIALS = "" Then 
				  strSQL = strSQL & "NULL)"
				Else       
				  strSQL = strSQL & "'" & sBADGE_INITIALS & "')"
				End If
			
				conDB.Execute strSQL, lngRecs, adCmdtext
			
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End If
			End If
		End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
'update ticket info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		If Len(Request.Form("GOTTICKETS")) Then 
			iBadgeNum = Request.Form("GOTTICKETS")
			
			strSQL = "select guest_id from badge where badge_num = " & iBadgeNum & " and guest_id in " & _
					 "(select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A')"
					 
			Set rsGuestID = conDB.Execute(strSQL, lngRecs, adCmdtext)
			
			iGuestID = rsGuestID("guest_id")
			
			Set rsGuestID = Nothing
			
			strSQL = "INSERT INTO log_ticket_activity (" & _
					 "acttimestamp, guest_id, pickup_name) VALUES ("
					 
			strSQL = strSQL & "GETDATE(), " & iGuestID & ", '" & ManageQuotes(sBADGE_PICKUP_NAME) & "')"
			
			conDB.Execute strSQL, lngRecs, adCmdtext 
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
'update program info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		strSQL = "select count(*) as badges from badge where guest_id in " & _
		         "(select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A')"
		
		Set rsNumBadges = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		iNumBadges = rsNumBadges("badges")
		
		Set rsNumBadges = Nothing
		
		For idx = 1 to iNumBadges
			
			If Len(Trim(Request("GOTPROGRAMS" & idx ))) Then
				Session("PROGRAM_COUNT") = Session("PROGRAM_COUNT") & "," & idx - 1 & "-" & Trim(Request.Form("GOTPROGRAMS" & idx))
'first see if the program count has been changed by comparing it to the hidden field
				If Trim(Request.Form("GOTPROGRAMS" & idx)) <> Trim(Request.Form("OLD_NUM_PROGRAMS" & idx)) Then
'got to get the guest_id. Badge_num is the same as the idx (that's the way it was set up in exhibitor.asp)
					strSQL = "select guest_id from badge where badge_num = " & idx & " and guest_id in " & _
							 "(select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A')"
					
					Set rsGuestID = conDB.Execute(strSQL, lngRecs, adCmdtext)
					
					iGuestID = rsGuestID("guest_id")
					
					Set rsGuestID = Nothing
					
'now update the log tables with the info
					strSQL = "INSERT INTO log_program_activity (" & _
							 "acttimestamp, num_programs, guest_id, pickup_name) VALUES ("
							 
					strSQL = strSQL & "GETDATE(), " & Trim(Request.Form("GOTPROGRAMS" & idx)) & ", " & iGuestID & ", '" & ManageQuotes(sBADGE_PICKUP_NAME) & "')"
					
					conDB.Execute strSQL, lngRecs, adCmdtext 
					
					If lngRecs <> 1 Then
						iserror = iserror + 1
					End if
				End If
			End If
		Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'log_badge_activity and print badges for this exhibitor based on the number of printbadge checkboxes checked
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'does the badge need to be printed
'the printbadge value is the badge_num
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Session("PRINT_ME") = cPRINTBADGE
		Session("PRINT_PROGRAM_RECEIPT") = cINCPROGRAMPAYMENT
		Session("NEEDSRECEIPT") = cINCPAYMENT
			
		If Len(Request.Form("PRINTBADGE")) or Len(Session("NEEDSRECEIPT")) or Len(Session("PRINT_PROGRAM_RECEIPT")) or Len(Session("PRINT_ME")) Then
		
		cPRINTBADGE = 1
			arBadgeNums = split(Trim(Request.Form("PRINTBADGE")), ",")
			
			For idx = LBound(arBadgeNums) to UBound(arBadgeNums)
'get the badge_id for the badge_num
				strSQL = "select badge_id from badge where guest_id in " & _
						 "(select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A') " & _
						 "and badge_num = " & arBadgeNums(idx)
						 
				Set rsBadgeID = conDB.Execute(strSQL, lngRecs, adCmdtext)
				
				iBadgeID = rsBadgeID("badge_id")
				
				Set rsBadgeID = Nothing
'now update the log tables with the info
				
				strSQL = "INSERT INTO log_badge_activity (" & _
						 "acttimestamp, badge_id, pickup_name) VALUES ("
						 
				strSQL = strSQL & "GETDATE(), " & iBadgeID & ", '" & ManageQuotes(sBADGE_PICKUP_NAME) & "')"
				
				conDB.Execute strSQL, lngRecs, adCmdtext 
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
			Next
		End If

		If iserror = 0 Then
			conDB.CommitTrans
		Else
			conDB.RollbackTrans
		End If

		conDB.Close
		
		Set conDB = Nothing		
	End If
	
		

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon

	Session("PRINT_ME") = " " & Session("PRINT_ME") & ","
	Session("NEEDSRECEIPT") = " " & Session("NEEDSRECEIPT") & ","
	Session("PRINT_PROGRAM_RECEIPT") = " " & Session("PRINT_PROGRAM_RECEIPT") & ","
	Session("PROGRAM_COUNT") = Session("PROGRAM_COUNT") & ","
	
	programtest = split(Session("PROGRAM_COUNT"),",")

	strSQL = "select e.company_name, eg.first_name, eg.last_name, eg.guest_id, e.preregistered, b.badge_num " & _
			 "from exhibitor e, exhibitor_guest eg, badge b " & _
			 "where e.exhibitor_id = eg.exhibitor_id " & _
			 "and eg.guest_id = b.guest_id " & _
			 "and eg.exhibitor_id = " & Session("EXHIBITOR_ID") & _
			 " and eg.status = 'A' "
			 
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)

Do While Not rs.EOF
		sAmount = "0.00"
		
		badge_string_num = " " & rs("badge_num") & ","
		
		intPrintBadge = instr(Session("PRINT_ME"), badge_string_num)
		
		intPrintBadgeReceipt = instr(Session("NEEDSRECEIPT"), badge_string_num)
		
		intPrintProgramReceipt = instr(Session("PRINT_PROGRAM_RECEIPT"), badge_string_num)
		
		'sName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 27)
		sName = rs("first_name") & "  " & rs("last_name")
		
		sCompany =  rs("company_name")
		
		If Len(sCompany) > 25 Then
'find the closest space before character 26, get it's number
			s = Mid(sCompany, 1, 26)  's = United States Air Force Ba
			
			if StrComp(Mid(s, 26, 1), Chr(32)) = 0 then
				For i = 26 To 1 Step -1
					t = StrComp(Mid(s, i, 1), Chr(32))
					
					If t = 0 Then
						sCompany1 = Mid(sCompany, 1, i)
						
						sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
						
						Exit For
					End If
				Next
			else
				For i = 31 To 1 Step -1
					t = StrComp(Mid(s, i, 1), Chr(32))
					
					If t = 0 Then
						sCompany1 = Mid(sCompany, 1, i)
						
						sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
						
						Exit For
					End If
				Next
			end if
'place all text before it in sCompany1, and all after it in sCompany2
		Else
			sCompany1 = sCompany
			
			sCompany2 = ""
		End If
		
		If rs("preregistered") = "Y" Then
			sPayType = "Preregistration"
		Else
			sPayType = "Registration"
		End If
		
		If instr(Session("NEEDSRECEIPT"), rs("badge_num")) Then
			sAmount = "135.00"
		Else
			sAmount = " 0.00"
		End If

		If intPrintBadge > 0 Then
			sBadgeString = "^XA" & _
							"^A0N,100,100^FO0,170^FB810,3,,C^FD" & sName & "^FS"
				
			If Len(sName) > 17 Then				
				If sCompany2 <> "" Then
					sBadgeString = sBadgeString & "^A0N,60,60^FO0,400^FB810,3,,C^FD" & sCompany1 & "" & sCompany2 & "^FS" & _
	"^A0N,40,40^FO0,600^FB810,3,,C^FDExhibitor^FS" & _
	"^XZ"					
				Else
					sBadgeString = sBadgeString & "^A0N,60,60^FO0,400^FB810,3,,C^FD" & sCompany1 & "" & sCompany2 & "^FS" & _
	"^A0N,40,40^FO0,500^FB810,3,,C^FDExhibitor^FS" & _
	"^XZ"		
				End If
			Else
				If sCompany2 <> "" Then
					sBadgeString = sBadgeString & "^A0N,60,60^FO0,300^FB810,3,,C^FD" & sCompany1 & "" & sCompany2 & "^FS" & _
	"^A0N,40,40^FO0,500^FB810,3,,C^FDExhibitor^FS" & _
	"^XZ"					
				Else
					sBadgeString = sBadgeString & "^A0N,60,60^FO0,300^FB810,3,,C^FD" & sCompany1 & "" & sCompany2 & "^FS" & _
	"^A0N,40,40^FO0,400^FB810,3,,C^FDExhibitor^FS" & _
	"^XZ"		
				End If
			End If
			
			printobject = printobject & "Set objFS = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine
			printobject = printobject & "Set objPrint = objFS.CreateTextFile(""LPT1:"", True)" & vbNewLine
			printobject = printobject & "objPrint.Write(""" & sBadgeString & """)" & vbNewLine
			printobject = printobject & "objPrint.Close" & vbNewLine
		End If

		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 then
			sReceiptString= "^XA" & _
"^A0N,40,40^FO40,200^FDREGISTRATION PAYMENT DETAILS:^FS" & _
"^A0N,40,40^FO40,250^FDPAY TYPE^FS" & _
"^A0N,40,40^FO350,250^FDDATE^FS" & _
"^A0N,40,40^FO650,250^FDAMOUNT^FS"
		End If
			
		If intPrintBadgeReceipt > 0 then
			sReceiptString = sReceiptString & "^A0N,40,40^FO40,320^FD" & sPayType & "^FS" & _
							"^A0N,40,40^FO320,320^FD" & Date & "^FS" & _
							"^A0N,40,40^FO650,320^FD" & sAmount & "^FS" 
		End If
			
		If intPrintProgramReceipt > 0 Then
			For i = 1 to 7
				programtester = instr(Left(programtest(i),1),rs("badge_num") - 1)
				
				If programtester > 0 then
					sProgramAmount = Right(programtest(i),Len(programtest(i)) - 2)
					
					sProgramAmount = 20 * sProgramAmount
										
					Exit For
				End If
			Next			
		End If
			
		If intPrintBadgeReceipt > 0 or intPrintProgramReceipt > 0 Then
			sReceiptString = sReceiptString & "^XZ"
			
			printobject = printobject & "Set objFS = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine
			printobject = printobject & "Set objPrint = objFS.CreateTextFile(""LPT1:"", True)" & vbNewLine
			printobject = printobject & "objPrint.Write(""" & sReceiptString & """)" & vbNewLine
			printobject = printobject & "objPrint.Close" & vbNewLine
		End If
		
					%>
			<html>
	<body>
		<script type="text/vbscript">
			
			<%= printobject%>
		</script>
<%
			printobject = null
		
		rs.movenext		
	Loop

	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing	

%>
		<script type="text/vbscript">
			
			window.location = "exhibitor.asp"
		</script>
	</body>
</html>
