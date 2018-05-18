<% @Language=VBScript %>
<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim sBadgeString
	dim sPayType
	dim s75
	dim s25
	dim sZero
	dim sFrom
	dim btype
	dim sNAME
	dim sZIP
	dim sCITY
	dim sSTATE
	dim sCOUNTRY
	dim sBADGETYPE
	dim sPAID
	dim sPREREG
	
	sNAME = Trim(Request("NAME"))
	sZIP = Trim(Request("ZIP"))
	sCITY = Trim(Request("CITY"))
	sSTATE = Trim(Request("STATES"))
	sCOUNTRY = Trim(Request("COUNTRY"))
	sBADGETYPE = Trim(Request("BADGETYPE"))
	sPAID = Trim(Request("PAID"))
	sPREREG = Trim(Request("PREREG"))
	sSchoolCompany = Trim(Request("SCHOOL_COMPANY"))

	
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
	
	s75 = "DLG"
	s25 = "S"
	sZero = "CF"

	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
	
	sNAME = Mid(sNAME, 1, 20)
	sFrom = Mid(sCITY & ", " & sSTATE, 1, 27)
'''		
'		sSchoolCompany = rs("school_company")
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
'''		

	
	If sCOUNTRY = "United States" Then
		sCOUNTRY = "&nbsp;"
	Else
		sCOUNTRY = Mid(sCOUNTRY, 1, 27)
	End If
	
	'If sBADGETYPE = "Director/Adult" Then
		'btype = "Director/Adult"
	'ElseIf sBADGETYPE = "Commercial" Then
		'btype = "Commercial"
	'ElseIf sBADGETYPE = "Guest" Then
		'btype = "Guest"
	'ElseIf sBADGETYPE = "College Student" Then
		'btype = "College Student"
	'ElseIf sBADGETYPE = "Clinician" Then
		'btype = "Clinician"
	'ElseIf sBADGETYPE = "Staff" Then
		'btype = "Staff"
	'End If
	
	If sPREREG = "Yes" Then
		sPayType = "Preregistration"
	Else
		sPayType = "Registration"
	End If
		
		sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:36px'><b>" & sName & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:20px'><b>" & sSC1 & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:20px'><b>" & sSC2 & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sFrom & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCountry & "</b></td></tr>" & _
									  "<tr><td colspan='3' align='center' style='font-size:8px'><b>&nbsp;</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:16px'><b>&nbsp;</b></td>" & _
									  "<td width='170' align='center' style='font-size:20px'><b>" & sBADGETYPE & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='2' align='left' style='font-size:16px'><b>" & sName & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:16px'><b>&nbsp;</b></td></tr>" & _
									  "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" & _
									  "<td width='80' align='right' style='font-size:16px'><b>$" & sPAID & "</b></td></tr>" & _
									  "<tr><td colspan='3' style='font-size:16px'>&nbsp;</td></tr>" & _
									  "<tr><td colspan='3' align='left' style='font-size:16px'><b>MIDWEST CLINIC PAYMENT DETAILS:</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:14px'><b>PAY TYPE</b></td>" & _
									  "<td width='170' align='center' style='font-size:14px'><b>DATE</b></td>" & _
									  "<td width='80' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>" & _
									  "<tr><td width='80' align='left' style='font-size:14px'><b>" & sPayType & "</b></td>" & _
									  "<td width='170' align='center' style='font-size:14px'><b>" & Date & "</b></td>" & _
									  "<td width='80' align='right' style='font-size:14px'><b>$" & sPAID & "</b></td></tr>"  
	
	sBadgeString = sBadgeString & "</table>"
%>
<html>
	<head>
		<style type="text/css">
body,td,th {
	font-family: Arial Narrow;
}
</style>
	</head>
	<!--<BODY onLoad="javascript:window.print();javascript:window.close();">-->
	<body onLoad="javascript:window.print();javascript:window.close();" onUnload="javascript:window.open('manual_attendee_badge.asp', 'manual_attendee_badge');">
		<% =sBadgeString %>
	</body>
</html>
