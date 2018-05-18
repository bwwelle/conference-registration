<% @Language=VBScript %>
<% Response.Buffer="true"%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%

Server.ScriptTimeout = 999

Function GetPRExhibitorBadges()
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from  badge b, exhibitor_guest eg, exhibitor e "
		strSQL = strSQL & "where eg.guest_id = b.guest_id " 
		strSQL = strSQL & "and e.exhibitor_id = eg.exhibitor_id "
		strSQL = strSQL & "and eg.status = 'A' " 
		strSQL = strSQL & "and b.badge_type = 'E' "
		strSQL = strSQL & "and e.preregistered = 'Y' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetPRExhibitorBadges = iCount
End Function

Function GetPRDirectorBadges()
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, registration r "
		strSQL = strSQL & "where b.registration_id = r.registration_id "
		strSQL = strSQL & "and b.badge_type = 'D' "
		strSQL = strSQL & "and r.preregistered = 'Y'"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetPRDirectorBadges = iCount
End Function

Function GetExhibitorBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from  badge b, log_badge_activity l, exhibitor_guest e "
		strSQL = strSQL & "where e.guest_id = b.guest_id "
		strSQL = strSQL & "and b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and e.status = 'A' "
		strSQL = strSQL & "and b.badge_type = 'E'"
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from  badge b, log_badge_activity l, exhibitor_guest e "
		strSQL = strSQL & "where e.guest_id = b.guest_id "
		strSQL = strSQL & "and b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and e.status = 'A' "
		strSQL = strSQL & "and b.badge_type = 'E')"

	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetExhibitorBadges = iCount
	

End Function

Function GetDirectorBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'D' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'D' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetDirectorBadges = iCount
End Function

Function GetGuestBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'G' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'G' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetGuestBadges = iCount
End Function

Function GetCommercialBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'L' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'L' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCommercialBadges = iCount
End Function

Function GetCollegeBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'S' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'S' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCollegeBadges = iCount
End Function

Function GetClinicianBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'C' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'C' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetClinicianBadges = iCount
End Function

Function GetStaffBadges(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select count(distinct b.badge_id) badges "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_id = l.badge_id "
		strSQL = strSQL & "and l.acttimestamp like 'Dec " & day & " 2010%' "
		strSQL = strSQL & "and b.badge_type = 'F' "
		strSQL = strSQL & "and b.badge_id not in "
		strSQL = strSQL & "(select b.badge_id "
		strSQL = strSQL & "from badge b, log_badge_activity l "
		strSQL = strSQL & "where b.badge_type = 'F' "
		strSQL = strSQL & "and l.acttimestamp < 'Dec " & day & " 2010' "
		strSQL = strSQL & "and b.badge_id = l.badge_id)"
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iCount = rs("badges")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetStaffBadges = iCount
End Function

Function GetCABadgePayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
				 "where payment_type = 'CA' and acttimestamp like 'Dec " & day & " 2010%' "
	
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCABadgePayment = iPayment 
End Function

Function GetCCBadgePayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
				 "where payment_type = 'CC' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCCBadgePayment = iPayment 
End Function


Function GetCHBadgePayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_badge_payment_activity " & _
				 "where payment_type = 'CH' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCHBadgePayment = iPayment 
End Function

Function GetCABoothPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
				 "where payment_type = 'CA' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCABoothPayment = iPayment 
End Function

Function GetCCBoothPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
				 "where payment_type = 'CC' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCCBoothPayment = iPayment 
End Function


Function GetCHBoothPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(amount_paid), 0) amount_paid from log_booth_payment_activity " & _
				 "where payment_type = 'CH' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCHBoothPayment = iPayment 
End Function

Function GetCAProgramPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(num_programs), 0) * 20 amount_paid from program " & _
				 "where payment_type = 'CA' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCAProgramPayment = iPayment 
End Function

Function GetCCProgramPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(num_programs), 0) * 20 amount_paid from program " & _
				 "where payment_type = 'CC' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCCProgramPayment = iPayment 
End Function

Function GetCHProgramPayment(day)
	Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Open strCon
		strSQL = "select isnull(sum(num_programs), 0) * 20 amount_paid from program " & _
				 "where payment_type = 'CH' and acttimestamp like 'Dec " & day & " 2010%' "
	Set rs = conDB.Execute(strSQL, lngRecs, adCmdtext)
	iPayment = rs("amount_paid")
	Set rs = Nothing
	conDB.Close
	Set conDB = Nothing
	GetCHProgramPayment = iPayment 
End Function

	dim conDB
	dim strSQL
	dim lngRecs
	dim rs
	dim sREGDATE
	dim sSearchDateReg
	dim sSearchDateAct
	dim sSearchDateAct2
	dim sPrereg
	dim iCountExhibitor
	dim iCountAdult
	dim iCountGuest
	dim iCountCommercial
	dim iCountCollege
	dim iCountClinician
	dim iCountStaff
	dim iSumBadgeCA
	dim iSumBadgeCC
	dim iSumBadgeCH
	dim iSumBoothCA
	dim iSumBoothCC
	dim iSumBoothCH
	dim iTotalCountBadges
	dim iTotalBadges
	dim iTotalBooths
	dim iTotalCash
	dim iTotalCheck
	dim iTotalCredit
	dim iTotal
	dim iMon
	dim iTue
	dim iWed
	dim iThu
	dim iFri
	dim iSat
	
	iMon = 13
	iTue = 14
	iWed = 15
	iThu = 16
	iFri = 17
	iSat = 18


'	Set conDB = Server.CreateObject("ADODB.Connection")
'	conDB.Open strCon
	sPrereg = ""
	sREGDATE = Trim(Request("REGDATE"))
	'sTEST = Trim(Request("TEST"))
		
	if sREGDATE = "Week To Date" Then
		sSearchDateAct = ">= 'Dec 12 2005'"
		sSearchDateAct2 = "<= 'Dec 11 2005'" 
	Elseif sREGDATE = "Week To Date Badge Details" Then
		Response.Redirect("report_details.asp")
	Elseif sREGDATE = "Preregistered" Then
		sPrereg = "Y"
	Elseif sREGDATE = "Exhibitor Report" Then
		Response.Redirect("report_exhibitors_present.asp")
	Elseif sREGDATE = "Attendee Payment Log" Then
		Response.Redirect("report_attendee_payment_log.asp")
		sSearchDateAct = "<= 'Dec 11 2005'"
		sSearchDateAct2 = ">= 'Dec 12 2005'"
	End If
	
'	If sPrereg = "Y" Then
	
		'get the counts for all the separate badge types printed for the selected date
		'for exhibitor
		'get the counts for all the separate badge types printed for the selected date

		iCountExhibitorMon = GetExhibitorBadges(iMon)
		iCountExhibitorTue = GetExhibitorBadges(iTue)
		iCountExhibitorWed = GetExhibitorBadges(iWed)
		iCountExhibitorThu = GetExhibitorBadges(iThu)
		iCountExhibitorFri = GetExhibitorBadges(iFri)
		iCountExhibitorSat = GetExhibitorBadges(iSat)
		iCountAdultMon = GetDirectorBadges(iMon)
		iCountAdultTue = GetDirectorBadges(iTue)
		iCountAdultWed = GetDirectorBadges(iWed)
		iCountAdultThu = GetDirectorBadges(iThu)
		iCountAdultFri = GetDirectorBadges(iFri)
		iCountAdultSat = GetDirectorBadges(iSat)
		iCountGuestMon = GetGuestBadges(iMon)
		iCountGuestTue = GetGuestBadges(iTue)
		iCountGuestWed = GetGuestBadges(iWed)
		iCountGuestThu = GetGuestBadges(iThu)
		iCountGuestFri = GetGuestBadges(iFri)
		iCountGuestSat = GetGuestBadges(iSat)
		iCountCommercialMon = GetCommercialBadges(iMon)
		iCountCommercialTue = GetCommercialBadges(iTue)
		iCountCommercialWed = GetCommercialBadges(iWed)
		iCountCommercialThu = GetCommercialBadges(iThu)
		iCountCommercialFri = GetCommercialBadges(iFri)
		iCountCommercialSat = GetCommercialBadges(iSat)
		iCountCollegeMon = GetCollegeBadges(iMon)
		iCountCollegeTue = GetCollegeBadges(iTue)
		iCountCollegeWed = GetCollegeBadges(iWed)
		iCountCollegeThu = GetCollegeBadges(iThu)
		iCountCollegeFri = GetCollegeBadges(iFri)
		iCountCollegeSat = GetCollegeBadges(iSat)
		iCountClinicianMon = GetClinicianBadges(iMon)
		iCountClinicianTue = GetClinicianBadges(iTue)
		iCountClinicianWed = GetClinicianBadges(iWed)
		iCountClinicianThu = GetClinicianBadges(iThu)
		iCountClinicianFri = GetClinicianBadges(iFri)
		iCountClinicianSat = GetClinicianBadges(iSat)
		iCountStaffMon = GetStaffBadges(iMon)
		iCountStaffTue = GetStaffBadges(iTue)
		iCountStaffWed = GetStaffBadges(iWed)
		iCountStaffThu = GetStaffBadges(iThu)
		iCountStaffFri = GetStaffBadges(iFri)
		iCountStaffSat = GetStaffBadges(iSat)
		
		iSumBadgeCAMon = GetCABadgePayment(iMon)
		iSumBadgeCATue = GetCABadgePayment(iTue)
		iSumBadgeCAWed = GetCABadgePayment(iWed)
		iSumBadgeCAThu = GetCABadgePayment(iThu)
		iSumBadgeCAFri = GetCABadgePayment(iFri)
		iSumBadgeCASat = GetCABadgePayment(iSat)
		
		iSumBadgeCCMon = GetCCBadgePayment(iMon)
		iSumBadgeCCTue = GetCCBadgePayment(iTue)
		iSumBadgeCCWed = GetCCBadgePayment(iWed)
		iSumBadgeCCThu = GetCCBadgePayment(iThu)
		iSumBadgeCCFri = GetCCBadgePayment(iFri)
		iSumBadgeCCSat = GetCCBadgePayment(iSat)
		
		iSumBadgeCHMon = GetCHBadgePayment(iMon)
		iSumBadgeCHTue = GetCHBadgePayment(iTue)
		iSumBadgeCHWed = GetCHBadgePayment(iWed)
		iSumBadgeCHThu = GetCHBadgePayment(iThu)
		iSumBadgeCHFri = GetCHBadgePayment(iFri)
		iSumBadgeCHSat = GetCHBadgePayment(iSat)
		
		iSumBoothCAMon = GetCABoothPayment(iMon) 
		iSumBoothCATue = GetCABoothPayment(iTue) 
		iSumBoothCAWed = GetCABoothPayment(iWed) 
		iSumBoothCAThu = GetCABoothPayment(iThu) 
		iSumBoothCAFri = GetCABoothPayment(iFri) 
		iSumBoothCASat = GetCABoothPayment(iSat)
		 
		iSumBoothCCMon = GetCCBoothPayment(iMon) 
		iSumBoothCCTue = GetCCBoothPayment(iTue) 
		iSumBoothCCWed = GetCCBoothPayment(iWed) 
		iSumBoothCCThu = GetCCBoothPayment(iThu) 
		iSumBoothCCFri = GetCCBoothPayment(iFri) 
		iSumBoothCCSat = GetCCBoothPayment(iSat)
		 
		iSumBoothCHMon = GetCHBoothPayment(iMon) 
		iSumBoothCHTue = GetCHBoothPayment(iTue) 
		iSumBoothCHWed = GetCHBoothPayment(iWed) 
		iSumBoothCHThu = GetCHBoothPayment(iThu) 
		iSumBoothCHFri = GetCHBoothPayment(iFri) 
		iSumBoothCHSat = GetCHBoothPayment(iSat) 
		
		iSumProgramCAMon = GetCAProgramPayment(iMon) 
		iSumProgramCATue = GetCAProgramPayment(iTue) 
		iSumProgramCAWed = GetCAProgramPayment(iWed) 
		iSumProgramCAThu = GetCAProgramPayment(iThu) 
		iSumProgramCAFri = GetCAProgramPayment(iFri) 
		iSumProgramCASat = GetCAProgramPayment(iSat)
		 
		iSumProgramCCMon = GetCCProgramPayment(iMon) 
		iSumProgramCCTue = GetCCProgramPayment(iTue) 
		iSumProgramCCWed = GetCCProgramPayment(iWed) 
		iSumProgramCCThu = GetCCProgramPayment(iThu) 
		iSumProgramCCFri = GetCCProgramPayment(iFri) 
		iSumProgramCCSat = GetCCProgramPayment(iSat)
		 
		iSumProgramCHMon = GetCHProgramPayment(iMon) 
		iSumProgramCHTue = GetCHProgramPayment(iTue) 
		iSumProgramCHWed = GetCHProgramPayment(iWed) 
		iSumProgramCHThu = GetCHProgramPayment(iThu) 
		iSumProgramCHFri = GetCHProgramPayment(iFri) 
		iSumProgramCHSat = GetCHProgramPayment(iSat) 
		
		
'	End If
	
	iTotalExhibitor = (iCountExhibitorMon + iCountExhibitorTue + iCountExhibitorWed + iCountExhibitorThu + iCountExhibitorFri + iCountExhibitorSat)
	iTotalAdult = (iCountAdultMon + iCountAdultTue + iCountAdultWed + iCountAdultThu + iCountAdultFri + iCountAdultSat)
	iTotalGuest = (iCountGuestMon + iCountGuestTue + iCountGuestWed + iCountGuestThu + iCountGuestFri + iCountGuestSat)	
	iTotalCommercial = (iCountCommercialMon + iCountCommercialTue + iCountCommercialWed + iCountCommercialThu + iCountCommercialFri + iCountCommercialSat)
	iTotalCollege = (iCountCollegeMon + iCountCollegeTue + iCountCollegeWed + iCountCollegeThu + iCountCollegeFri + iCountCollegeSat)
	iTotalClinician = (iCountClinicianMon + iCountClinicianTue + iCountClinicianWed + iCountClinicianThu + iCountClinicianFri + iCountClinicianSat)	
	iTotalStaff = (iCountStaffMon + iCountStaffTue + iCountStaffWed + iCountStaffThu + iCountStaffFri + iCountStaffSat)
	
	iTotalMon=(iCountExhibitorMon + iCountAdultMon + iCountGuestMon + iCountCommercialMon + iCountCollegeMon + iCountClinicianMon + iCountStaffMon)
	iTotalTue=(iCountExhibitorTue + iCountAdultTue + iCountGuestTue + iCountCommercialTue + iCountCollegeTue + iCountClinicianTue + iCountStaffTue)
	iTotalWed=(iCountExhibitorWed + iCountAdultWed + iCountGuestWed + iCountCommercialWed + iCountCollegeWed + iCountClinicianWed + iCountStaffWed)
	iTotalThu=(iCountExhibitorThu + iCountAdultThu + iCountGuestThu + iCountCommercialThu + iCountCollegeThu + iCountClinicianThu + iCountStaffThu)
	iTotalFri=(iCountExhibitorFri + iCountAdultFri + iCountGuestFri + iCountCommercialFri + iCountCollegeFri + iCountClinicianFri + iCountStaffFri)
	iTotalSat=(iCountExhibitorSat + iCountAdultSat + iCountGuestSat + iCountCommercialSat + iCountCollegeSat + iCountClinicianSat + iCountStaffSat)
	iTotalBadges=(iTotalMon + iTotalTue + iTotalWed + iTotalThu + iTotalFri + iTotalSat)
	
	iTotalSumBooths=(iSumBoothCAMon + iSumBoothChMon + iSumBoothCcMon + iSumBoothCAtue + iSumBoothChtue + iSumBoothCctue + iSumBoothCAwed + iSumBoothChwed + iSumBoothCcwed + iSumBoothCAthu + iSumBoothChthu + iSumBoothCcthu + iSumBoothCAfri + iSumBoothChfri + iSumBoothCcfri + iSumBoothCAsat + iSumBoothChsat + iSumBoothCcsat)
	iTotalSumBadges=(iSumBadgeCAMon + iSumBadgeChMon + iSumBadgeCcMon + iSumBadgeCAtue + iSumBadgeChtue + iSumBadgeCctue + iSumBadgeCAwed + iSumBadgeChwed + iSumBadgeCcwed + iSumBadgeCAthu + iSumBadgeChthu + iSumBadgeCcthu + iSumBadgeCAfri + iSumBadgeChfri + iSumBadgeCcfri + iSumBadgeCAsat + iSumBadgeChsat + iSumBadgeCcsat)
	iTotalSumPrograms=(iSumProgramCAMon + iSumProgramChMon + iSumProgramCcMon + iSumProgramCAtue + iSumProgramChtue + iSumProgramCctue + iSumProgramCAwed + iSumProgramChwed + iSumProgramCcwed + iSumProgramCAthu + iSumProgramChthu + iSumProgramCcthu + iSumProgramCAfri + iSumProgramChfri + iSumProgramCcfri + iSumProgramCAsat + iSumProgramChsat + iSumProgramCcsat)

	iTotalProgramsMon = (iSumProgramCAMon + iSumProgramCCMon + iSumProgramCHMon)/20
	iTotalProgramsTue = (iSumProgramCATue + iSumProgramCCTue + iSumProgramCHTue)/20
	iTotalProgramsWed = (iSumProgramCAWed + iSumProgramCCWed + iSumProgramCHWed)/20
	iTotalProgramsThu = (iSumProgramCAThu + iSumProgramCCThu + iSumProgramCHThu)/20
	iTotalProgramsFri = (iSumProgramCAFri + iSumProgramCCFri + iSumProgramCHFri)/20
	iTotalProgramsSat = (iSumProgramCASat + iSumProgramCCSat + iSumProgramCHSat)/20

	
	iTotalSumMon=(iSumBadgeCAmon + iSumBoothCAMon + iSumProgramCAmon + iSumBadgeChmon + iSumBoothChmon + iSumProgramChmon + iSumBadgeCcMon + iSumBoothCcMon + iSumProgramCCmon)
	iTotalSumTue=(iSumBadgeCAtue + iSumBoothCAtue + iSumProgramCAtue + iSumBadgeChtue + iSumBoothChtue + iSumProgramChtue + iSumBadgeCctue + iSumBoothCctue + iSumProgramCCtue)
	iTotalSumWed=(iSumBadgeCAwed + iSumBoothCAwed + iSumProgramCAwed + iSumBadgeChwed + iSumBoothChwed + iSumProgramChwed + iSumBadgeCcwed + iSumBoothCcwed + iSumProgramCCwed)
	iTotalSumThu=(iSumBadgeCAthu + iSumBoothCAthu + iSumProgramCAthu + iSumBadgeChthu + iSumBoothChthu + iSumProgramChthu + iSumBadgeCcthu + iSumBoothCcthu + iSumProgramCCthu)
	iTotalSumFri=(iSumBadgeCAfri + iSumBoothCAfri + iSumProgramCAfri + iSumBadgeChfri + iSumBoothChfri + iSumProgramChfri + iSumBadgeCcfri + iSumBoothCcfri + iSumProgramCCfri)
	iTotalSumSat=(iSumBadgeCAsat + iSumBoothCAsat + iSumProgramCAsat + iSumBadgeChsat + iSumBoothChsat + iSumProgramChsat + iSumBadgeCcsat + iSumBoothCcsat + iSumProgramCCsat)
	iTotalSum=(iTotalSumMon + iTotalSumTue + iTotalSumWed + iTotalSumThu + iTotalSumFri + iTotalSumSat)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration - Report View</title>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body>
		<table border="1" width="100%" bordercolor="gray" cellspacing="0">
			<tr bgcolor="#AAD5FF">
				<td colspan="25" align="center">
					<b>
						Badge Counts
					</b>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td colspan="5" align="center">
					<b>
						Badge Type
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Monday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Tuesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Wednesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Thursday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Friday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Saturday
					</b>
				</td>
				<td colspan="2" align="center">
					<b>
						Total
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						Exhibitor
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountExhibitorSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalExhibitor%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td colspan="5" align="left">
					<b>
						Adult/Director
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountAdultSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalAdult%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						Guest
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountGuestSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalGuest%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td colspan="5" align="left">
					<b>
						Commercial
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCommercialSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalCommercial%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						College Student
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountCollegeSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalCollege%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td colspan="5" align="left">
					<b>
						Clinician
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountClinicianSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalClinician%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						Staff
					</b>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffMon%>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffTue%>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffWed%>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffThu%>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffFri%>
				</td>
				<td colspan="3" align="right">
					<%=iCountStaffSat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalStaff%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="5" align="left">
					<b>
						Total Badges
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalMon%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalTue%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalWed%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalThu%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalFri%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSat%>
					</b>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalBadges%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="25">&nbsp;
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="25" align="center">
					<b>
						Program Counts
					</b>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td colspan="5" align="center">&nbsp;
				</td>
				<td colspan="3" align="center">
					<b>
						Monday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Tuesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Wednesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Thursday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Friday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Saturday
					</b>
				</td>
				<td colspan="2" align="center">
					<b>
						Total
					</b>
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="5" align="left">
					<b>
						Total Programs
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsMon%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsTue%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsWed%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsThu%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsFri%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalProgramsSat%>
					</b>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=(iTotalProgramsMon + iTotalProgramsTue + iTotalProgramsWed + iTotalProgramsThu + iTotalProgramsFri + iTotalProgramsSat)%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="25">&nbsp;
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="25" align="center">
					<b>
						OnSite Payments
					</b>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td colspan="5" align="center">&nbsp;
				</td>
				<td colspan="3" align="center">
					<b>
						Monday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Tuesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Wednesday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Thursday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Friday
					</b>
				</td>
				<td colspan="3" align="center">
					<b>
						Saturday
					</b>
				</td>
				<td colspan="2" align="center">
					<b>
						Total
					</b>
				</td>
			</tr>
			<tr bgcolor="#DFEFFF">
				<td colspan="5" align="left">
					<b>
						Payment Type
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ca
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Ch
					</b>
				</td>
				<td width="4%" align="center">
					<b>
						Cc
					</b>
				</td>
				<td colspan="2" align="center">&nbsp;
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						Badges
					</b>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCctue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcsat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalSumBadges%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#F4FAFF">
				<td colspan="5" align="left">
					<b>
						Booths
					</b>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCcMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCctue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCcwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCcthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCcfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCAsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothChsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBoothCcsat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalSumBooths%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="left">
					<b>
						Programs
					</b>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCcMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCctue%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCcwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCcthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCcfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCAsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramChsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumProgramCcsat%>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalSumPrograms%>
					</b>
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="5" align="left">
					<b>
						Total By Payment Type
					</b>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAmon + iSumBoothCAMon + iSumProgramCAMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCHMon + iSumBoothCHMon + iSumProgramCHMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCCMon + iSumBoothCCMon + iSumProgramCCMon%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCATue + iSumBoothCATue + iSumProgramCATue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChtue + iSumBoothCHTue + iSumProgramCATue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCctue + iSumBoothCctue + iSumProgramCCtue%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAwed + iSumBoothCAwed + iSumProgramCAwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChwed + iSumBoothChwed + iSumProgramChwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcwed + iSumBoothCcwed + iSumProgramCcwed%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAthu + iSumBoothCAthu + iSumProgramCAthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChthu + iSumBoothChthu + iSumProgramChthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcthu + iSumBoothCcthu + iSumProgramCcthu%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAfri + iSumBoothCAfri + iSumProgramCAfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChfri + iSumBoothChfri + iSumProgramChfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcfri + iSumBoothCcfri + iSumProgramCcfri%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCAsat + iSumBoothCAsat + iSumProgramCAsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeChsat + iSumBoothChsat + iSumProgramChsat%>
				</td>
				<td width="4%" align="right">
					<%=iSumBadgeCcsat + iSumBoothCcsat + iSumProgramCcsat%>
				</td>
				<td colspan="2" align="right">
					<b>
						&nbsp;
					</b>
				</td>
			</tr>
			<tr bgcolor="#AAD5FF">
				<td colspan="5" align="left">
					<b>
						Total By Day
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumMon%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumTue%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumWed%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumThu%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumFri%>
					</b>
				</td>
				<td colspan="3" align="right">
					<b>
						<%=iTotalSumSat%>
					</b>
				</td>
				<td colspan="2" align="right">
					<b>
						<%=iTotalSum%>
					</b>
				</td>
			</tr>
			<tr>
				<td colspan="25" align="center">
					<a href="report_select.asp">
						Return to Select Reports
					</a>
				</td>
			</tr>
		</table>
	</body>
</html>
