<% @Language=VBScript %>
<% response.buffer="true" %>
<% Server.ScriptTimeout = 999 %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sEXHIBITOR_ID
	dim sBOOTH_NUMBERS
	dim badge_cost_due
	dim sADDL_BADGE_COUNT
	dim iComp
	dim iNComp
	dim iAddlComp
	dim iAddlNonComp
	dim rs
	dim newBadgeCount
	dim lngRecs
	dim strSQL
	dim conDB
	dim lngCommas
	dim idx
	dim iserror
	dim numBadges
	dim arBadges
	dim arBadgeCount()
	dim numCount
	dim newCompCount
	dim newNonCompCount
	dim rsMaxBadgeNum
	dim newBadgeNum   				
	
	numBadges = ((Session("NUM_BOOTHS") * 6) + 20)
	
	set conDB = CreateObject("ADODB.Connection") 
	conDB.open strCon
	
	iserror = 0
	
	numCount = 0
	
'first delete so we can reorder the badge_nums
'the badge array for deleting is a list of the guest_ids
	If Len(Request.form("chkBadge")) Then
		arBadges = split(Trim(Request.form("chkBadge")), ",")
		
		For idx = LBound(arBadges) to UBound(arBadges)
				strSQL = "Update exhibitor_guest set status = 'D' where exhibitor_id = " & Session("EXHIBITOR_ID") & " and guest_id = " & arBadges(idx)
				
				conDB.Execute strSQL, lngRecs, adCmdtext
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
				
				numCount = numCount + 1 'this will be how many were deleted
		Next 
				
'now renumber the badge_nums - and can't use .recordcount for whatever stupid reason...
		strSQL = "select count(*) as guest_count from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A' "
		
		Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		newBadgeCount = rs("guest_count")
		
		Set rs = Nothing
		
		strSQL = "select guest_id from badge where guest_id in " & _
		         "(select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A') " & _
				 "order by badge_num"
		
		Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		redim arBadgeCount(newBadgeCount)
		
		For idx = 1 to newBadgeCount
			If not rs.EOF then
				arBadgeCount(idx) = rs("guest_id")
				
				rs.Movenext
			End If
		Next

		Set rs = Nothing
		
'now update the badge_nums
		For idx = 1 to newBadgeCount
			If arBadgeCount(idx) <> null then
				strSQL = "update badge set badge_num = " & idx & " where guest_id = " & arBadgeCount(idx)
				
				conDB.Execute strSQL, lngRecs, adCmdtext
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
			End If
		Next
		
'update comp/noncomp in exhibitor table for deleted badges
'first get all badges
		strSQL = "select count(*) as guest_count from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A' "
		
		Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		newBadgeCount = rs("guest_count")
		
		Set rs = Nothing
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'if there are more than the allowed comps
'''''''''''''''''''''''''''''''''''''''''''''''''''''
		If newBadgeCount >= (Session("NUM_BOOTHS") * 6) Then
			newBadgeCount = newBadgeCount - (Session("NUM_BOOTHS") * 6)
			
			strSQL = "update exhibitor set addl_badges_non_comp = " & newBadgeCount & _
					 ", badge_amount_due = " & newBadgeCount * 135 & _
					 " where exhibitor_id = " & Session("EXHIBITOR_ID")
			
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		Else
			strSQL = "update exhibitor set addl_badges_non_comp = 0, badge_amount_due = 0 where exhibitor_id = " & Session("EXHIBITOR_ID")
			
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if		
		End If 
	
	
	Else	
'save the new and updated badges
		For idx = 1 to numBadges
			If Trim(Request("FIRSTNAME" & idx )) <> "" Then
'hasn't been saved yet
				If Trim(Request("GUESTID" & idx )) = "" Then
'is not in tables so insert
					strSQL = "Insert into exhibitor_guest values (" & Session("EXHIBITOR_ID") & ", '" & ManageQuotes(Trim(Request("FIRSTNAME" & idx ))) & "', '" & ManageQuotes(Trim(Request("LASTNAME" & idx ))) & "', 'A')"
					
					conDB.Execute strSQL, lngRecs, adCmdtext
					
					If lngRecs <> 1 Then
						iserror = iserror + 1
					End if

'now get the guest_id (cuz remember it was "")
					strSQL = "select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & _
							 " and first_name = '" & ManageQuotes(Trim(Request("FIRSTNAME" & idx ))) & _
							 "' and last_name = '" & ManageQuotes(Trim(Request("LASTNAME" & idx ))) & _
							 "' and status = 'A'"							 
					
					Set rsGuestID = conDB.Execute(strSQL, lngRecs, adCmdtext)
					
					newGuestID = rsGuestID("guest_id")
					
					Set rsGuestID = Nothing
					
					set rsMaxBadgeNum = CreateObject("ADODB.Recordset")
					
'is it comp - based on the booth count
					strSQL = "select max(badge_num) as ""Badge_Num"" from badge where guest_id in (select guest_id from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A')"
											
					rsMaxBadgeNum.open strSQL,strCon,1,1
				    
					If IsNull(rsMaxBadgeNum("Badge_Num")) or (rsMaxBadgeNum.cachesize < 1) then
						newBadgeNum = 0						
					else
						newBadgeNum = rsMaxBadgeNum("Badge_Num")
					end if
					
					newBadgeNum = newBadgeNum + 1
					
					rsMaxBadgeNum.Close
					
					Set rsMaxBadgeNum = Nothing
					
					strSQL = "Insert into badge (badge_type, guest_id, badge_num) "
strSQL = strSQL & "values ('E', " & newGuestID & ", " &  newBadgeNum  & ")"
					
					conDB.Execute strSQL, lngRecs, adCmdtext
							
					If lngRecs <> 1 Then
						iserror = iserror + 1
					End if
				ElseIf Trim(Request("GUESTID" & idx )) <> "" Then
'the guest is being edited so update
					strSQL = "update exhibitor_guest set first_name = '" & ManageQuotes(Trim(Request("FIRSTNAME" & idx ))) & _
							 "', last_name = '" & ManageQuotes(Trim(Request("LASTNAME" & idx ))) & "' " & _
							 "where exhibitor_id = " & Session("EXHIBITOR_ID") & _
							 " and guest_id = " & Trim(Request("GUESTID" & idx ))
					
					conDB.Execute strSQL, lngRecs, adCmdtext
					
					If lngRecs <> 1 Then
						iserror = iserror + 1
					End if
				End If
			End If
		Next 
		
'now update the exhibitor table 

'update comp/noncomp in exhibitor table for inserted/updated badges
'first get all badges
		strSQL = "select count(*) as guest_count from exhibitor_guest where exhibitor_id = " & Session("EXHIBITOR_ID") & " and status = 'A' "
		
		Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		newBadgeCount = rs("guest_count")
		
		Set rs = Nothing
		
'check that allowed comps are >= NUM_BOOTHS * 6, if not, skip to insert
		If newBadgeCount >= (Session("NUM_BOOTHS") * 6) Then
'subtract allowed comps Session("NUM_BOOTHS") * 6
			newBadgeCount = newBadgeCount - (Session("NUM_BOOTHS") * 6)
			
			strSQL = "update exhibitor set addl_badges_non_comp = " & newBadgeCount & ", badge_amount_due = " & (newBadgeCount * 135) & _
					 " where exhibitor_id = " & Session("EXHIBITOR_ID")
			
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		Else
			strSQL = "update exhibitor set addl_badges_non_comp = 0, badge_amount_due = 0 where exhibitor_id = " & Session("EXHIBITOR_ID")
			
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		End If
	End If
	
	If iserror = 0 Then
	End If
	
	conDB.Close
	
	Set conDB = Nothing
	
	Response.Redirect("exhibitor.asp")
%>
