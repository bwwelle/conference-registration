<% @Language=VBScript %>

<!--#include virtual="/includes/dbfunctions.asp" -->
<!--#include virtual="/includes/dbglobals.asp" -->
<!--#include virtual="/includes/adovbs.inc"-->

<%
	dim sBadgeString
	dim lngRecs
	dim strSQL
	dim conDB
	dim rs
	dim sPayType
	dim sAmount
	dim sName
	dim sCompany
	dim sCompany1
	dim sCompany2
	dim s
	dim t
	dim i

	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon

	strSQL = "select e.company_name, eg.first_name, eg.last_name, eg.guest_id, e.preregistered, b.badge_num " & _
			 "from exhibitor e, exhibitor_guest eg, badge b " & _
			 "where e.exhibitor_id = eg.exhibitor_id " & _
			 "and eg.guest_id = b.guest_id " & _
			 "and eg.exhibitor_id = " & Session("EXHIBITOR_ID") & _
			 " and eg.status = 'A' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	sBadgeString = "<table width='330' align='center' border=0 cellspacing='0'>"
	Do While Not rs.EOF
		If instr(Session("PRINT_ME"), rs("badge_num")) Then
		
			sName = Mid(rs("first_name") & "  " & rs("last_name"), 1, 20)
			sCompany =  rs("company_name")
            If Len(sCompany) > 25 Then
                'find the closest space before character 26, get it's number
                s = Mid(sCompany, 1, 26)  's = United States Air Force Ba
                For i = 26 To 1 Step -1
                    t = StrComp(Mid(s, i, 1), Chr(32))
                    If t = 0 Then
                        sCompany1 = Mid(sCompany, 1, i)
                        sCompany2 = Mid(sCompany, i + 1, Len(sCompany))
                        Exit For
                    End If
                Next
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
				sAmount = "75.00"
			Else
				sAmount = " 0.00"
			End If
				
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:36px'><b>" & sName & "</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:10px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany1 & "</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='center' style='font-size:26px'><b>" & sCompany2 & "</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:8px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:16px'><b>" & rs("guest_id") & "</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:20px'><b>EXHIBITOR</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:20px'><b>&nbsp;</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:22px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='2' align='left' style='font-size:16px'><b>" & sName & "</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:16px'><b>" & rs("guest_id") & "</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='2' align='left' style='font-size:16px'><b>TOTAL PAID:</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:16px'><b>$" & sAmount & "</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' style='font-size:16px'>&nbsp;</td></tr>" 
			sBadgeString = sBadgeString & "<tr><td colspan='3' align='left' style='font-size:16px'><b>MIDWEST CLINIC PAYMENT DETAILS:</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>PAYMENT TYPE</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>DATE</b></td>"
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>AMOUNT</b></td></tr>" 
			sBadgeString = sBadgeString & "<tr><td width='110' align='left' style='font-size:14px'><b>" & sPayType & "</b></td>"
			sBadgeString = sBadgeString & "<td width='110' align='center' style='font-size:14px'><b>" & Date & "</b></td>" 
			sBadgeString = sBadgeString & "<td width='110' align='right' style='font-size:14px'><b>$" & sAmount & "</b></td></tr>"  
		End If
		rs.movenext
	Loop
	sBadgeString = sBadgeString & "</table>"
	Set rs = Nothing
	
	conDB.Close
	Set conDB = Nothing
%>

<HTML>
<HEAD>
<style type="text/css">
body,td,th {
	font-family: Arial Narrow;
}
</style>
</HEAD>
<BODY onLoad="javascript:window.print();javascript:window.close();">
<form name="printform">

<% =sBadgeString %> 

</form>
</BODY>
</HTML>