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
			sCountry = "&nbsp;"
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
ElseIf rs("badge_type") = "BS" Then
			btype = "BOARD SPOUSE"
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
		ElseIf instr(s25, rs("badge_type")) and sPayType = "Preregistration" Then
			sAmount = "50.00"
		ElseIf instr(s25, rs("badge_type")) and sPayType = "Registration" Then
			sAmount = "75.00"
		End If
		
		sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:36px'><b>" & sBadgeName & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:20px'><b>" & sSC1 & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:20px'><b>" & sSC2 & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sBadgeFrom & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCountry & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:8px'><b>&nbsp;</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:16px'><b>" & rs("registration_id") & "</b></td>" & _
									  "<td width='170' align='center' style='font-size:20px'><b>" & btype & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='2' align='left' style='font-size:16px'><b>" & sName & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:16px'><b>" & rs("registration_id") & "</b></td></tr>" & _
									  "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" & _
									  "<td width='80' align='right' style='font-size:16px'><b>$" & sAmount & "</b></td></tr>" & _
									  "<tr><td colspan='3' style='font-size:16px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' align='left' style='font-size:16px'><b>REGISTRATION PAYMENT DETAILS:</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:14px'><b>PAY TYPE</b></td>" & _
									  "<td width='170' align='center' style='font-size:14px'><b>DATE</b></td>" & _
									  "<td width='80' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:14px'><b>" & sPayType & "</b></td>" & _
									  "<td width='170' align='center' style='font-size:14px'><b>" & Date & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:14px'><b>$" & sAmount & "</b></td></tr>"  
	End If
	
	sBadgeString = sBadgeString & "</table>"
	
	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" encoding="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
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
