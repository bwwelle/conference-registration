<% @Language=VBScript %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim sBadgeString
	dim sReceiptString
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
		End If
		
		If rs("preregistered") = "Y" Then
			sPayType = "Preregistration"
		Else
			sPayType = "Registration"
		End If
		
		If rs("complimentary") = "Y" Or instr(sZero, rs("badge_type")) Then
			sAmount = " 0.00"
		ElseIf instr(s75, rs("badge_type")) and sPayType = "Preregistration" Then
			sAmount = "90.00"
		ElseIf instr(s75, rs("badge_type")) and sPayType = "Registration" Then
			sAmount = "100.00"
		ElseIf instr(s25, rs("badge_type")) Then
			sAmount = "45.00"
		End If
		
		sBadgeString = "^XA^A0N,100,100^FO0,70^FB816,1,,C^FD" & sBadgeName & "^FS^A0N,50,50^FO0,170^FB816,1,,C^FD" & sSC1 & "^FS" & _
"^A0N,70,70^FO0,270^FB816,1,,C^FD" & sBadgeFrom & "^FS" & _
"^A0N,70,70^FO0,340^FB816,1,,C^FD" & sCountry & "^FS" & _
"^A0N,50,50^FO10,420^FD" & rs("registration_id") & "^FS" & _
"^A0N,50,50^FO0,480^FB816,1,,C^FD" & btype & "^FS" & _
"^XZ"

		sReceiptString= "^XA" & _
"^A0N,40,40^FO40,30^FD" & sName & "^FS" & _
"^A0N,40,40^FO680,30^FD" & rs("registration_id") & "^FS" & _
"^A0N,40,40^FO40,80^FDTOTAL PAID:^FS" & _
"^A0N,40,40^FO680,80^FD" & sAmount & "^FS" & _
"^A0N,40,40^FO40,200^FDREGISTRATION PAYMENT DETAILS:^FS" & _
"^A0N,40,40^FO40,250^FDPAY TYPE^FS" & _
"^A0N,40,40^FO350,250^FDDATE^FS" & _
"^A0N,40,40^FO650,250^FDAMOUNT^FS" & _
"^A0N,40,40^FO40,320^FD" & sPayType & "^FS" & _
"^A0N,40,40^FO320,320^FD" & Date & "^FS" & _
"^A0N,40,40^FO650,320^FD" & sAmount & "^FS" & _
"^XZ"
	End If
	Set rs = Nothing
	
	conDB.Close
	
	Set conDB = Nothing

	
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objPrint = objFS.CreateTextFile("LPT1:", True)
objPrint.Write(sBadgeString)
objPrint.Close

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objPrint = objFS.CreateTextFile("LPT1:", True)
objPrint.Write(sReceiptString)
objPrint.Close
%>

<html>
	<head>
	</head>
	<body onLoad = "javascript:window.close();">
	</body>
</html>
