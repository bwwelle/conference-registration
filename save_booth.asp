<% @Language=VBScript %>
<% response.buffer="true" %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%	
	dim sEXHIBITOR_ID
	dim sBOOTH_NUMBER
	dim sBOOTH_DELETES
	dim sTOT_BOOTH_COUNT
	dim lngRecs
	dim strSQL
	dim conDB
	dim lngCommas
	dim arBooths
	dim idx
	dim sHall
	dim iserror
	dim num
	
	sBOOTH_NUMBER = Trim(Request("BOOTH_NUMBER"))
	
	sBOOTH_DELETES = Trim(Request.form("chkBooth"))
	
	sEXHIBITOR_ID = Session("EXHIBITOR_ID")
	
'''''add booths to the table'''''
	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon
	
	conDB.BeginTrans
	
	iserror = 0
	
	Session("BOOTH_IS_USED") = ""

	If Len(Request("BOOTH_NUMBER")) Then
'first find out if the booth is already being used - if not, save new booth number
		strSQL = "select count(*) booths from booth where booth_number = '" & sBOOTH_NUMBER & "' and status = 'A'"

		Set rsBoothCount = conDB.Execute(strSQL, lngRecs, adCmdtext)
		
		num = rsBoothCount("booths")
		
		Set rsBoothCount = Nothing
		
		If num = 0 Then
			sHall = "Other"
			
			Dim alphatest
			alphatest = 0
			
			if right(trim(sBooth_Number), 1) = "a" then
				sBooth_Number = mid(trim(sBooth_Number), 1, len(trim(sBooth_Number))-1)
				
				alphatest = 1
			End if
			
			If (trim(sBOOTH_NUMBER) >= 2000) And (trim(sBOOTH_NUMBER) <= 2307) Then
				sHall = "Northeast Exhibit Hall"
			Elseif (trim(sBOOTH_NUMBER) >= 1300) And (trim(sBOOTH_NUMBER) <= 1832) Then
				sHall = "Northwest Exhibit Hall"
			Elseif (trim(sBOOTH_NUMBER) >= 700) And (trim(sBOOTH_NUMBER) <= 1208) Then
				sHall = "Southwest Exhibit Hall"
			Elseif (trim(sBOOTH_NUMBER) >= 101) And (trim(sBOOTH_NUMBER) <= 621) Then
				sHall = "Southeast Exhibit Hall"
			Elseif (trim(sBOOTH_NUMBER) <= 68) Then
				sHall = "Continental Ballroom"
			End If
			
			if alphatest = 1 then
				sBooth_Number = trim(sBooth_Number) & "a"
			end if
			
			strSQL = "Insert into Booth values (" & sEXHIBITOR_ID & ", " & trim(sBOOTH_NUMBER) & ", '" & sHall & "', 'A')"
			
			conDB.Execute strSQL, lngRecs, adCmdtext
			
			If lngRecs <> 1 Then
				iserror = iserror + 1
			End if
		Else
			Session("BOOTH_IS_USED") = "Booth number " & sBOOTH_NUMBER & " is in use."	
		End If
	End If
	
	If Len(Request.form("chkBooth")) Then
		arBooths = split(Trim(Request.form("chkBooth")), ",")

		For idx = LBound(arBooths) to UBound(arBooths)
				strSQL = "Update Booth set status = 'D' where exhibitor_id = " & sEXHIBITOR_ID & " and booth_number = '" & arBooths(idx) & "'"
				
				conDB.Execute strSQL, lngRecs, adCmdtext
				
				If lngRecs <> 1 Then
					iserror = iserror + 1
				End if
		Next 
	End If	
	
	
'''''update badge/booth info for main exhibitor screen to display'''''
	strSQL = "select count(booth_id) booth_id from booth where exhibitor_id = " & sEXHIBITOR_ID & " and status = 'A'"
	
	Set rsBoothCount = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	sTOT_BOOTH_COUNT = rsBoothCount("booth_id")
	
	Set rsBoothCount = Nothing
	
	strSQL = "Update exhibitor set booth_amount_due = " & (CInt(sTOT_BOOTH_COUNT) * 990) & _
	         " where exhibitor_id = " & sEXHIBITOR_ID 
	
	conDB.Execute strSQL, lngRecs, adCmdtext
	
	If lngRecs <> 1 Then
		iserror = iserror + 1
	End if
	
	If iserror = 0 Then
		conDB.CommitTrans
	End If
	
	conDB.Close
	
	Set conDB = Nothing
	
	Response.Redirect("exhibitor_booth_modify.asp")
%>
