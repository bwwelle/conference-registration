<% @Language=VBScript %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<%
	dim conDB  
	dim strSQL
	dim lngRecs
	dim rsZip
	
	Session("LOOKUP_CITY") = ""
	Session("LOOKUP_STATE") = ""
	Session("LOOKUP_ZIP") = ""
		
	sLOOKUPTHISZIP = Trim(Request("LOOKUPTHISZIP"))

	Set conDB = Server.CreateObject("ADODB.Connection")
	
	conDB.Open strCon

	strSQL = "SELECT city, state FROM zipcodeworld1 " & _
			 "WHERE zip_code = '" & Trim(ManageQuotes(sLOOKUPTHISZIP)) & "' "
			 
	Set rsZip = conDB.Execute(strSQL, lngRecs, adCmdtext)
	
	If Not rsZip.EOF Then
		Session("LOOKUP_CITY") = trim(rsZip("city"))
		
		Session("LOOKUP_STATE") = trim(rsZip("state"))
		
		Session("LOOKUP_ZIP") = sLOOKUPTHISZIP
	End If
	
	Set rsZip = Nothing
	
	conDB.Close
	
	Set conDB = Nothing	
	
	Session("SEARCH_FOR_ZIP") = "false"
	
	Response.Redirect("registration.asp")	
%>
