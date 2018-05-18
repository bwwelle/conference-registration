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
	dim sCOMPANY
	dim sCompany1
	dim sCompany2
	dim sPAID
	dim sPREREG
	dim s
	dim t
	dim i
	
	sNAME = Trim(Request("NAME"))
	sCOMPANY = Trim(Request("COMPANY"))
	sPAID = Trim(Request("PAID"))
	sPREREG = Trim(Request("PREREG"))

	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
	
	sNAME = Mid(sNAME, 1, 20)
	If Len(sCOMPANY) > 25 Then
		'find the closest space before character 26, get it's number
		s = Mid(sCOMPANY, 1, 26)  's = United States Air Force Ba
		For i = 26 To 1 Step -1
			t = StrComp(Mid(s, i, 1), Chr(32))
			If t = 0 Then
				sCompany1 = Mid(sCOMPANY, 1, i)
				sCompany2 = Mid(sCOMPANY, i + 1, Len(sCOMPANY))
				Exit For
			End If
		Next
		'place all text before it in sCompany1, and all after it in sCompany2

	Else
		sCompany1 = sCOMPANY
		sCompany2 = ""
	End If
	
	If sPREREG = "Yes" Then
		sPayType = "Preregistration"
	Else
		sPayType = "Registration"
	End If
		
	sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' align='center' style='font-size:36px'><b>" & sNAME & "</b></td></tr>" & _
								  "<tr><td colspan='3' style='font-size:10px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany1 & "</b></td></tr>" & _
								  "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany2 & "</b></td></tr>" & _
								  "<tr><td colspan='3' align='center' style='font-size:8px'><b>&nbsp;</b></td></tr>" & _
								  "<tr><td width='70' align='left' style='font-size:16px'><b>&nbsp;</b></td>" & _
								  "<td width='190' align='center' style='font-size:20px'><b>Exhibitor</b></td>" & _
								  "<td width='70' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='2' align='left' style='font-size:16px'><b>" & sNAME & "</b></td>" & _
								  "<td width='110' align='right' style='font-size:16px'><b>&nbsp;</b></td></tr>" & _
								  "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" & _
								  "<td width='110' align='right' style='font-size:16px'><b>$" & sPAID & "</b></td></tr>" & _
								  "<tr><td colspan='3' style='font-size:16px'>&nbsp;</td></tr>" & _
								  "<tr><td colspan='3' align='left' style='font-size:16px'><b>MIDWEST CLINIC PAYMENT DETAILS:</b></td></tr>" & _
								  "<tr><td width='110' align='left' style='font-size:14px'><b>PAYMENT TYPE</b></td>" & _
								  "<td width='110' align='center' style='font-size:14px'><b>DATE</b></td>" & _
								  "<td width='110' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>" & _
								  "<tr><td width='110' align='left' style='font-size:14px'><b>Preregistration</b></td>" & _
								  "<td width='110' align='center' style='font-size:14px'><b>" & Date & "</b></td>" & _
								  "<td width='110' align='right' style='font-size:14px'><b>$" & sPAID & "</b></td></tr>"  
	
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
	<body onLoad="javascript:window.print();javascript:window.close();" onUnload="javascript:window.open('manual_exhibitor_badge.asp', 'manual_exhibitor_badge');">
		<form name="printform">
			<% =sBadgeString %>
		</form>
	</body>
</html>
