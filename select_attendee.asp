<% @Language = VBSCRIPT %>
<% Response.Buffer="true"%>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->

<% Server.ScriptTimeout = 999%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>Registration Select Attendee</title>
		<script language="javascript">
			function KeyDownHandler(btn)
			{
// process only the Enter key
				if (event.keyCode == 13)
				{
// cancel the default submit
					event.returnValue=false;
					
					event.cancel = true;
					
// submit the form by programmatically clicking the specified button
					btn.click();
				}
			}
		
			function setfocus() 
			{
			  document.select_attendee.SEARCH_LAST_NAME.focus();
			}	
		</script>
		<style type="text/css">
			body,td,th 
			{
				font-family: Arial, Helvetica, sans-serif;
				
				font-size: x-small;
			}
		</style>
	</head>
	<body onLoad="setfocus()">
<!-- BEGIN CONTENT -->
		<form name="select_attendee" method=post action="select_attendee.asp">
			<table border=0>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							<input type="submit" name="NEW" value="Add New Attendee">
						</div>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							--or--
						</div>
					</td>
				</tr>
<!-- SEARCH BY NAME -->
				<tr>
					<td width=300>
						<div align="right">
							Search By Attendee Last Name
						</div>
					</td>
					<td width=100>
						<div align="left">
							<input type="text" id="SEARCH_LAST_NAME" name="SEARCH_LAST_NAME" size="50" value="" onKeyDown= "KeyDownHandler(SEARCH)">
						</div>
					</td>
					<td width=100>
						<div align="left">
							<input type="submit" name="SEARCH" id="SEARCH" value="Search">
						</div>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<%			
							dim rsNames
							dim indexName
							dim conDB
							dim strSQL
							dim indexCountry
							dim lngRecs
							
							Set conDB = Server.CreateObject("ADODB.Connection")
							
							conDB.Open strCon
							
							Session("ATTENDEE_LIST") = ""
							
							Session("SEARCH_FOR_ZIP") = "true"
						
							If Len(Request.form("NEW")) Then 
								Session("SEARCH_FOR_ZIP") = "true"
							
								Response.Redirect("registration.asp")								
							ElseIf Len(Request.form("SEARCH")) Then
								sCOUNTRY = Trim(Request("COUNTRY"))
								
								If Trim(Request("States")) = null then
									sSTATES = ""
								Else
									sSTATES	= Trim(Request("STATES"))
								End if
								
								if sSTATES= "Alabama" then sSTATESABBR= "AL" end if
								if sSTATES= "Alaska" then sSTATESABBR= "AK" end if
								if sSTATES= "Alberta" then sSTATESABBR= "AB" end if
								if sSTATES= "Arizona" then sSTATESABBR= "AZ" end if
								if sSTATES= "Arkansas" then sSTATESABBR= "AR" end if
								if sSTATES= "British Columbia" then sSTATESABBR= "BC" end if
								if sSTATES= "California" then sSTATESABBR= "CA" end if
								if sSTATES= "Colorado" then sSTATESABBR= "CO" end if
								if sSTATES= "Connecticut" then sSTATESABBR= "CT" end if
								if sSTATES= "Delaware" then sSTATESABBR= "DE" end if
								if sSTATES= "Florida" then sSTATESABBR= "FL" end if
								if sSTATES= "Georgia" then sSTATESABBR= "GA" end if
								if sSTATES= "Hawaii" then sSTATESABBR= "HI" end if
								if sSTATES= "Idaho" then sSTATESABBR= "ID" end if
								if sSTATES= "Illinois" then sSTATESABBR= "IL" end if
								if sSTATES= "Indiana" then sSTATESABBR= "IN" end if
								if sSTATES= "Iowa" then	STATES= "IA" end if
								if sSTATES= "Kansas" then sSTATESABBR= "KS" end if
								if sSTATES= "Kentucky" then sSTATESABBR= "KY" end if
								if sSTATES= "Louisiana" then sSTATESABBR= "LA" end if
								if sSTATES= "Maine" then sSTATESABBR= "ME" end if
								if sSTATES= "Manitoba" then sSTATESABBR= "MB" end if
								if sSTATES= "Maryland" then sSTATESABBR= "MD" end if
								if sSTATES= "Massachusetts" then sSTATESABBR= "MA" end if
								if sSTATES= "Michigan" then sSTATESABBR= "MI" end if
								if sSTATES= "Minnesota" then sSTATESABBR= "MN" end if
								if sSTATES= "Mississippi" then sSTATESABBR= "MS" end if
								if sSTATES= "Missouri" then sSTATESABBR= "MO" end if
								if sSTATES= "Montana" then sSTATESABBR= "MT" end if
								if sSTATES= "Nebraska" then sSTATESABBR= "NE" end if
								if sSTATES= "Nevada" then sSTATESABBR= "NV" end if
								if sSTATES= "New Brunswick" then sSTATESABBR= "NB" end if
								if sSTATES= "New Hampshire" then sSTATESABBR= "NH" end if
								if sSTATES= "New Jersey" then sSTATESABBR= "NJ" end if
								if sSTATES= "New Mexico" then sSTATESABBR= "NM" end if
								if sSTATES= "New South Wales" then sSTATESABBR= "NSW" end if
								if sSTATES= "New York" then sSTATESABBR= "NY" end if
								if sSTATES= "Newfoundland" then sSTATESABBR= "NL" end if
								if sSTATES= "North Carolina" then sSTATESABBR= "NC" end if
								if sSTATES= "North Dakota" then sSTATESABBR= "ND" end if
								if sSTATES= "Northern Territory" then sSTATESABBR= "NT" end if
								if sSTATES= "Northwest Territories" then sSTATESABBR= "NT" end if
								if sSTATES= "Nova Scotia" then sSTATESABBR= "NS" end if
								if sSTATES= "Nunavut" then sSTATESABBR= "NU" end if
								if sSTATES= "Ohio" then	STATES= "OH" end if
								if sSTATES= "Oklahoma" then sSTATESABBR= "OK" end if
								if sSTATES= "Ontario" then sSTATESABBR= "ON" end if
								if sSTATES= "Oregon" then sSTATESABBR= "OR" end if
								if sSTATES= "Pennsylvania" then sSTATESABBR= "PA" end if
								if sSTATES= "Prince Edward Island" then sSTATESABBR= "PE" end if
								if sSTATES= "Quebec" then sSTATESABBR= "QC" end if
								if sSTATES= "Queensland" then sSTATESABBR= "QLD" end if
								if sSTATES= "Rhode Island" then sSTATESABBR= "RI" end if
								if sSTATES= "Saskatchewan" then sSTATESABBR= "SK" end if
								if sSTATES= "South Australia" then sSTATESABBR= "SA" end if
								if sSTATES= "South Carolina" then sSTATESABBR= "SC" end if
								if sSTATES= "South Dakota" then sSTATESABBR= "SD" end if
								if sSTATES= "Tasmania" then sSTATESABBR= "TAS" end if
								if sSTATES= "Tennessee" then sSTATESABBR= "TN" end if
								if sSTATES= "Texas" then sSTATESABBR= "TX" end if
								if sSTATES= "Utah" then sSTATESABBR= "UT" end if
								if sSTATES= "Vermont" then sSTATESABBR= "VT" end if
								if sSTATES= "Victoria" then sSTATESABBR= "VIC" end if
								if sSTATES= "Virginia" then sSTATESABBR= "VA" end if
								if sSTATES= "Washington" then sSTATESABBR= "WA" end if
								if sSTATES= "Washington DC" then sSTATESABBR= "DC" end if
								if sSTATES= "West Virginia" then sSTATESABBR= "WV" end if
								if sSTATES= "Western Australia ST" then sSTATESABBR= "W AUST" end if
								if sSTATES= "Wisconsin" then sSTATESABBR= "WI" end if
								if sSTATES= "Wyoming" then sSTATESABBR= "WY" end if
								if sSTATES= "Yukon Territory" then sSTATESABBR= "YT" end if
						
								sAddCriteria = ""
								
								If sCOUNTRY <> "" and sSTATESABBR <> "" Then
									sAddCriteria = " badge_country = '" & sCOUNTRY & "' AND badge_state = '" & sSTATESABBR & "' "
								ElseIf sCOUNTRY <> "" and sSTATESABBR = "" Then
									sAddCriteria = " badge_country = '" & sCOUNTRY & "' "
								ElseIf sCOUNTRY = "" and sSTATESABBR <> "" Then
									sAddCriteria = " badge_state = '" & sSTATESABBR & "' "
								End If
						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' display the list returned from the search
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sSearchName
sSearchName = Replace(Trim(Request("SEARCH_LAST_NAME")),"'","''")
								If sAddCriteria = "" then		
									'strSQL = "SELECT ' ' pre, last_name, first_name, city, state, registration_id FROM registration " & _
									 '"WHERE last_name like '" & Trim(Request("SEARCH_LAST_NAME")) & "%' " & _
									 '"UNION " & _
									 '"SELECT '*' pre, badge_last_name last_name, badge_first_name first_name, badge_city city, badge_state state, registration_id FROM badge " & _
									 '"WHERE badge_last_name like '" & Trim(Request("SEARCH_LAST_NAME")) & "%' " & _
									 '"AND len(badge_last_name) > 0 " & _
									 '"ORDER BY last_name, first_name"
									 
									 strSQL = "SELECT badge_last_name last_name, badge_first_name first_name, badge_city city, badge_state state, registration_id FROM badge " & _
									 "WHERE badge_last_name like '" & sSearchName & "%' " & _
									 "AND len(badge_last_name) > 0 " & _
									 "ORDER BY last_name, first_name"
								Else
									'strSQL = "SELECT ' ' pre, last_name, first_name, city, state, registration_id FROM registration " & _
									 '"WHERE "& sAddCriteria & _
									 '"ORDER BY last_name, first_name"
									 '"SELECT '*' pre, badge_last_name last_name, badge_first_name first_name, badge_city city, badge_state state, registration_id FROM badge " & _
									 '"WHERE badge_last_name like '" & Trim(Request("SEARCH_LAST_NAME")) & "%' " & _
									 '"AND len(badge_last_name) > 0 " & sAddCriteria & _
									 '"ORDER BY last_name, first_name"
									 
									 strSQL = "SELECT badge_last_name last_name, badge_first_name first_name, badge_city city, badge_state state, badge_country country, registration_id FROM badge " & _
									 "WHERE badge_last_name like '" & sSearchName & "%' " & _
									 "AND len(badge_last_name) > 0 AND " & sAddCriteria & _
									 "ORDER BY last_name, first_name"
								End if
								
								Set rsNames = conDB.Execute(strSQL, lngRecs, adCmdtext)
						
'display the list returned from the search
								Response.Write("<select name=""ATTENDEE_LIST"" size=""20"" ONCHANGE=""document.select_attendee.submit(); return false;"">")
								
								Do While Not rsNames.EOF 
									'badgeNote = ""
									
									'if rsNames("pre") = "*" then
										'badgeNote = "**"
									'end if
									
									strSQL = "SELECT l.registration_id "
									strSQL = strSQL & "FROM log_badge_activity l WHERE "
									strSQL = strSQL & "l.registration_id = " & rsNames("registration_id")
						
									Set rsBadgePrint = conDB.Execute(strSQL, lngRecs, adCmdtext)
									
									printNote = ""
									
									If lngRecs = -1 then
										printNote = "^"
									End if
												
									'Response.Write ("<option>" & badgeNote & "  " & rsNames("last_name") & " | " & rsNames("first_name") & " | " & rsNames("city") &  " | " & rsNames("state") & printNote & "</div></option>")
									
									Response.Write ("<option>" & rsNames("last_name") & " | " & rsNames("first_name") & " | " & rsNames("city") &  " | " & rsNames("state") & printNote & "</div></option>")
																		
									rsNames.MoveNext
								Loop
								
								Response.Write("</div></select>")
							ElseIf Len(Request.form("SUBMIT")) Then 
								Session("SEARCH_FOR_ZIP") = "false"
								
								Response.Redirect("registration.asp")							
							ElseIf Len(Request.form("ATTENDEE_LIST")) Then 
								Session("ATTENDEE_LIST") = Request.form("ATTENDEE_LIST")
								
								Session("SEARCH_FOR_ZIP") = "false"
										
								Response.Redirect("registration.asp")						
							End if
						
							Set rsNames = Nothing
							
							conDB.Close
							
							Set conDB = Nothing
						%>
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">
						<div align="left">
							--advanced search--
						</div>
					</td>
				</tr>
<!-- Search by Country -->
				<tr>
					<td width=300>
						<div align="right">
							Select Country
						</div>
					</td>
					<td width=100>
						<select name="COUNTRY"  onChange="SEARCH.click()">
							<option value="<%=COUNTRY%>">
							</option>
							<option <% if COUNTRY="United States" then response.write (" SELECTED ")%>>
								United States
							</option>
							<option <% if COUNTRY="Albania" then response.write (" SELECTED ")%>>
								Albania
							</option>
							<option <% if COUNTRY="Algeria" then response.write (" SELECTED ")%>>
								Algeria
							</option>
							<option <% if COUNTRY="American Samoa" then response.write (" SELECTED ")%>>
								American Samoa
							</option>
							<option <% if COUNTRY="Andorra" then response.write (" SELECTED ")%>>
								Andorra
							</option>
							<option <% if COUNTRY="Angola" then response.write (" SELECTED ")%>>
								Angola
							</option>
							<option <% if COUNTRY="Anguilla" then response.write (" SELECTED ")%>>
								Anguilla
							</option>
							<option <% if COUNTRY="Antarctica" then response.write (" SELECTED ")%>>
								Antarctica
							</option>
							<option <% if COUNTRY="Antigua And Barbuda" then response.write (" SELECTED ")%>>
								Antigua And Barbuda
							</option>
							<option <% if COUNTRY="Argentina" then response.write (" SELECTED ")%>>
								Argentina
							</option>
							<option <% if COUNTRY="Armenia" then response.write (" SELECTED ")%>>
								Armenia
							</option>
							<option <% if COUNTRY="Aruba" then response.write (" SELECTED ")%>>
								Aruba
							</option>
							<option <% if COUNTRY="Australia" then response.write (" SELECTED ")%>>
								Australia
							</option>
							<option <% if COUNTRY="Austria" then response.write (" SELECTED ")%>>
								Austria
							</option>
							<option <% if COUNTRY="Azerbaijan" then response.write (" SELECTED ")%>>
								Azerbaijan
							</option>
							<option <% if COUNTRY="Bahamas" then response.write (" SELECTED ")%>>
								Bahamas
							</option>
							<option <% if COUNTRY="Bahrain" then response.write (" SELECTED ")%>>
								Bahrain
							</option>
							<option <% if COUNTRY="Bangladesh" then response.write (" SELECTED ")%>>
								Bangladesh
							</option>
							<option <% if COUNTRY="Barbados" then response.write (" SELECTED ")%>>
								Barbados
							</option>
							<option <% if COUNTRY="Belarus" then response.write (" SELECTED ")%>>
								Belarus
							</option>
							<option <% if COUNTRY="Belgium" then response.write (" SELECTED ")%>>
								Belgium
							</option>
							<option <% if COUNTRY="Belize" then response.write (" SELECTED ")%>>
								Belize
							</option>
							<option <% if COUNTRY="Benin" then response.write (" SELECTED ")%>>
								Benin
							</option>
							<option <% if COUNTRY="Bermuda" then response.write (" SELECTED ")%>>
								Bermuda
							</option>
							<option <% if COUNTRY="Bhutan" then response.write (" SELECTED ")%>>
								Bhutan
							</option>
							<option <% if COUNTRY="Bolivia" then response.write (" SELECTED ")%>>
								Bolivia
							</option>
							<option <% if COUNTRY="Bosnia and Herzegovina" then response.write (" SELECTED ")%>>
								Bosnia and Herzegovina
							</option>
							<option <% if COUNTRY="Botswana" then response.write (" SELECTED ")%>>
								Botswana
							</option>
							<option <% if COUNTRY="Bouvet Island" then response.write (" SELECTED ")%>>
								Bouvet Island
							</option>
							<option <% if COUNTRY="Brazil" then response.write (" SELECTED ")%>>
								Brazil
							</option>
							<option <% if COUNTRY="British Indian Ocean Territory" then response.write (" SELECTED ")%>>
								British Indian Ocean Territory
							</option>
							<option <% if COUNTRY="Brunei Darussalam" then response.write (" SELECTED ")%>>
								Brunei Darussalam
							</option>
							<option <% if COUNTRY="Bulgaria" then response.write (" SELECTED ")%>>
								Bulgaria
							</option>
							<option <% if COUNTRY="Burkina Faso" then response.write (" SELECTED ")%>>
								Burkina Faso
							</option>
							<option <% if COUNTRY="Burma" then response.write (" SELECTED ")%>>
								Burma
							</option>
							<option <% if COUNTRY="Burundi" then response.write (" SELECTED ")%>>
								Burundi
							</option>
							<option <% if COUNTRY="Cambodia" then response.write (" SELECTED ")%>>
								Cambodia
							</option>
							<option <% if COUNTRY="Cameroon" then response.write (" SELECTED ")%>>
								Cameroon
							</option>
							<option <% if COUNTRY="Canada" then response.write (" SELECTED ")%>>
								Canada
							</option>
							<option <% if COUNTRY="Cape Verde" then response.write (" SELECTED ")%>>
								Cape Verde
							</option>
							<option <% if COUNTRY="Cayman Islands" then response.write (" SELECTED ")%>>
								Cayman Islands
							</option>
							<option <% if COUNTRY="Central African Republic" then response.write (" SELECTED ")%>>
								Central African Republic
							</option>
							<option <% if COUNTRY="Chad" then response.write (" SELECTED ")%>>
								Chad
							</option>
							<option <% if COUNTRY="Chile" then response.write (" SELECTED ")%>>
								Chile
							</option>
							<option <% if COUNTRY="China" then response.write (" SELECTED ")%>>
								China
							</option>
							<option <% if COUNTRY="Christmas Island" then response.write (" SELECTED ")%>>
								Christmas Island
							</option>
							<option <% if COUNTRY="Cocos (Keeling) Islands" then response.write (" SELECTED ")%>>
								Cocos (Keeling) Islands
							</option>
							<option <% if COUNTRY="Colombia" then response.write (" SELECTED ")%>>
								Colombia
							</option>
							<option <% if COUNTRY="Comoros" then response.write (" SELECTED ")%>>
								Comoros
							</option>
							<option <% if COUNTRY="Congo" then response.write (" SELECTED ")%>>
								Congo
							</option>
							<option <% if COUNTRY="Congo, the Democratic Republic of the" then response.write (" SELECTED ")%>>
								Congo, the Democratic Republic of the
							</option>
							<option <% if COUNTRY="Cook Islands" then response.write (" SELECTED ")%>>
								Cook Islands
							</option>
							<option <% if COUNTRY="Costa Rica" then response.write (" SELECTED ")%>>
								Costa Rica
							</option>
							<option <% if COUNTRY="Cote d'Ivoire" then response.write (" SELECTED ")%>>
								Cote d'Ivoire
							</option>
							<option <% if COUNTRY="Croatia" then response.write (" SELECTED ")%>>
								Croatia
							</option>
							<option <% if COUNTRY="Cyprus" then response.write (" SELECTED ")%>>
								Cyprus
							</option>
							<option <% if COUNTRY="Czech Republic" then response.write (" SELECTED ")%>>
								Czech Republic
							</option>
							<option <% if COUNTRY="Denmark" then response.write (" SELECTED ")%>>
								Denmark
							</option>
							<option <% if COUNTRY="Djibouti" then response.write (" SELECTED ")%>>
								Djibouti
							</option>
							<option <% if COUNTRY="Dominica" then response.write (" SELECTED ")%>>
								Dominica
							</option>
							<option <% if COUNTRY="Dominican Republic" then response.write (" SELECTED ")%>>
								Dominican Republic
							</option>
							<option <% if COUNTRY="East Timor" then response.write (" SELECTED ")%>>
								East Timor
							</option>
							<option <% if COUNTRY="Ecuador" then response.write (" SELECTED ")%>>
								Ecuador
							</option>
							<option <% if COUNTRY="Egypt" then response.write (" SELECTED ")%>>
								Egypt
							</option>
							<option <% if COUNTRY="El Salvador" then response.write (" SELECTED ")%>>
								El Salvador
							</option>
							<option <% if COUNTRY="England" then response.write (" SELECTED ")%>>
								England
							</option>
							<option <% if COUNTRY="Equatorial Guinea" then response.write (" SELECTED ")%>>
								Equatorial Guinea
							</option>
							<option <% if COUNTRY="Eritrea" then response.write (" SELECTED ")%>>
								Eritrea
							</option>
							<option <% if COUNTRY="Espana" then response.write (" SELECTED ")%>>
								Espana
							</option>
							<option <% if COUNTRY="Estonia" then response.write (" SELECTED ")%>>
								Estonia
							</option>
							<option <% if COUNTRY="Ethiopia" then response.write (" SELECTED ")%>>
								Ethiopia
							</option>
							<option <% if COUNTRY="Falkland Islands" then response.write (" SELECTED ")%>>
								Falkland Islands
							</option>
							<option <% if COUNTRY="Faroe Islands" then response.write (" SELECTED ")%>>
								Faroe Islands
							</option>
							<option <% if COUNTRY="Fiji" then response.write (" SELECTED ")%>>
								Fiji
							</option>
							<option <% if COUNTRY="Finland" then response.write (" SELECTED ")%>>
								Finland
							</option>
							<option <% if COUNTRY="France" then response.write (" SELECTED ")%>>
								France
							</option>
							<option <% if COUNTRY="French Guiana" then response.write (" SELECTED ")%>>
								French Guiana
							</option>
							<option <% if COUNTRY="French Polynesia" then response.write (" SELECTED ")%>>
								French Polynesia
							</option>
							<option <% if COUNTRY="French Southern Territories" then response.write (" SELECTED ")%>>
								French Southern Territories
							</option>
							<option <% if COUNTRY="Gabon" then response.write (" SELECTED ")%>>
								Gabon
							</option>
							<option <% if COUNTRY="Gambia" then response.write (" SELECTED ")%>>
								Gambia
							</option>
							<option <% if COUNTRY="Georgia" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if COUNTRY="Germany" then response.write (" SELECTED ")%>>
								Germany
							</option>
							<option <% if COUNTRY="Ghana" then response.write (" SELECTED ")%>>
								Ghana
							</option>
							<option <% if COUNTRY="Gibraltar" then response.write (" SELECTED ")%>>
								Gibraltar
							</option>
							<option <% if COUNTRY="Great Britain" then response.write (" SELECTED ")%>>
								Great Britain
							</option>
							<option <% if COUNTRY="Greece" then response.write (" SELECTED ")%>>
								Greece
							</option>
							<option <% if COUNTRY="Greenland" then response.write (" SELECTED ")%>>
								Greenland
							</option>
							<option <% if COUNTRY="Grenada" then response.write (" SELECTED ")%>>
								Grenada
							</option>
							<option <% if COUNTRY="Guadeloupe" then response.write (" SELECTED ")%>>
								Guadeloupe
							</option>
							<option <% if COUNTRY="Guam" then response.write (" SELECTED ")%>>
								Guam
							</option>
							<option <% if COUNTRY="Guatemala" then response.write (" SELECTED ")%>>
								Guatemala
							</option>
							<option <% if COUNTRY="Guinea" then response.write (" SELECTED ")%>>
								Guinea
							</option>
							<option <% if COUNTRY="Guinea-Bissau" then response.write (" SELECTED ")%>>
								Guinea-Bissau
							</option>
							<option <% if COUNTRY="Guyana" then response.write (" SELECTED ")%>>
								Guyana
							</option>
							<option <% if COUNTRY="Haiti" then response.write (" SELECTED ")%>>
								Haiti
							</option>
							<option <% if COUNTRY="Heard and Mc Donald Islands" then response.write (" SELECTED ")%>>
								Heard and Mc Donald Islands
							</option>
							<option <% if COUNTRY="Honduras" then response.write (" SELECTED ")%>>
								Honduras
							</option>
							<option <% if COUNTRY="Hong Kong" then response.write (" SELECTED ")%>>
								Hong Kong
							</option>
							<option <% if COUNTRY="Hungary" then response.write (" SELECTED ")%>>
								Hungary
							</option>
							<option <% if COUNTRY="Iceland" then response.write (" SELECTED ")%>>
								Iceland
							</option>
							<option <% if COUNTRY="India" then response.write (" SELECTED ")%>>
								India
							</option>
							<option <% if COUNTRY="Indonesia" then response.write (" SELECTED ")%>>
								Indonesia
							</option>
							<option <% if COUNTRY="Ireland" then response.write (" SELECTED ")%>>
								Ireland
							</option>
							<option <% if COUNTRY="Israel" then response.write (" SELECTED ")%>>
								Israel
							</option>
							<option <% if COUNTRY="Italy" then response.write (" SELECTED ")%>>
								Italy
							</option>
							<option <% if COUNTRY="Jamaica" then response.write (" SELECTED ")%>>
								Jamaica
							</option>
							<option <% if COUNTRY="Japan" then response.write (" SELECTED ")%>>
								Japan
							</option>
							<option <% if COUNTRY="Jordan" then response.write (" SELECTED ")%>>
								Jordan
							</option>
							<option <% if COUNTRY="Kazakhstan" then response.write (" SELECTED ")%>>
								Kazakhstan
							</option>
							<option <% if COUNTRY="Kenya" then response.write (" SELECTED ")%>>
								Kenya
							</option>
							<option <% if COUNTRY="Kiribati" then response.write (" SELECTED ")%>>
								Kiribati
							</option>
							<option <% if COUNTRY="Korea (North)" then response.write (" SELECTED ")%>>
								Korea(North)
							</option>
							<option <% if COUNTRY="Korea, Republic of" then response.write (" SELECTED ")%>>
								Korea, Republic of
							</option>
							<option <% if COUNTRY="Korea (South)" then response.write (" SELECTED ")%>>
								Korea (South)
							</option>
							<option <% if COUNTRY="Kuwait" then response.write (" SELECTED ")%>>
								Kuwait
							</option>
							<option <% if COUNTRY="Kyrgyzstan" then response.write (" SELECTED ")%>>
								Kyrgyzstan
							</option>
							<option <% if COUNTRY="Lao People's Democratic Republic" then response.write (" SELECTED ")%>>
								Lao People's Democratic Republic
							</option>
							<option <% if COUNTRY="Latvia" then response.write (" SELECTED ")%>>
								Latvia
							</option>
							<option <% if COUNTRY="Lebanon" then response.write (" SELECTED ")%>>
								Lebanon
							</option>
							<option <% if COUNTRY="Lesotho" then response.write (" SELECTED ")%>>
								Lesotho
							</option>
							<option <% if COUNTRY="Liberia" then response.write (" SELECTED ")%>>
								Liberia
							</option>
							<option <% if COUNTRY="Liechtenstein" then response.write (" SELECTED ")%>>
								Liechtenstein
							</option>
							<option <% if COUNTRY="Lithuania" then response.write (" SELECTED ")%>>
								Lithuania
							</option>
							<option <% if COUNTRY="Luxembourg" then response.write (" SELECTED ")%>>
								Luxembourg
							</option>
							<option <% if COUNTRY="Macau" then response.write (" SELECTED ")%>>
								Macau
							</option>
							<option <% if COUNTRY="Macedonia" then response.write (" SELECTED ")%>>
								Macedonia
							</option>
							<option <% if COUNTRY="Madagascar" then response.write (" SELECTED ")%>>
								Madagascar
							</option>
							<option <% if COUNTRY="Malawi" then response.write (" SELECTED ")%>>
								Malawi
							</option>
							<option <% if COUNTRY="Malaysia" then response.write (" SELECTED ")%>>
								Malaysia
							</option>
							<option <% if COUNTRY="Maldives" then response.write (" SELECTED ")%>>
								Maldives
							</option>
							<option <% if COUNTRY="Mali" then response.write (" SELECTED ")%>>
								Mali
							</option>
							<option <% if COUNTRY="Malta" then response.write (" SELECTED ")%>>
								Malta
							</option>
							<option <% if COUNTRY="Marshall Islands" then response.write (" SELECTED ")%>>
								Marshall Islands
							</option>
							<option <% if COUNTRY="Martinique" then response.write (" SELECTED ")%>>
								Martinique
							</option>
							<option <% if COUNTRY="Mauritania" then response.write (" SELECTED ")%>>
								Mauritania
							</option>
							<option <% if COUNTRY="Mauritius" then response.write (" SELECTED ")%>>
								Mauritius
							</option>
							<option <% if COUNTRY="Mayotte" then response.write (" SELECTED ")%>>
								Mayotte
							</option>
							<option <% if COUNTRY="Mexico" then response.write (" SELECTED ")%>>
								Mexico
							</option>
							<option <% if COUNTRY="Micronesia, Federated States of" then response.write (" SELECTED ")%>>
								Micronesia, Federated States of
							</option>
							<option <% if COUNTRY="Moldova, Republic of" then response.write (" SELECTED ")%>>
								Moldova, Republic of
							</option>
							<option <% if COUNTRY="Monaco" then response.write (" SELECTED ")%>>
								Monaco
							</option>
							<option <% if COUNTRY="Mongolia" then response.write (" SELECTED ")%>>
								Mongolia
							</option>
							<option <% if COUNTRY="Montserrat" then response.write (" SELECTED ")%>>
								Montserrat
							</option>
							<option <% if COUNTRY="Morocco" then response.write (" SELECTED ")%>>
								Morocco
							</option>
							<option <% if COUNTRY="Mozambique" then response.write (" SELECTED ")%>>
								Mozambique
							</option>
							<option <% if COUNTRY="Myanmar" then response.write (" SELECTED ")%>>
								Myanmar
							</option>
							<option <% if COUNTRY="Namibia" then response.write (" SELECTED ")%>>
								Namibia
							</option>
							<option <% if COUNTRY="Nauru" then response.write (" SELECTED ")%>>
								Nauru
							</option>
							<option <% if COUNTRY="Nepal" then response.write (" SELECTED ")%>>
								Nepal
							</option>
							<option <% if COUNTRY="Netherlands" then response.write (" SELECTED ")%>>
								Netherlands
							</option>
							<option <% if COUNTRY="Netherlands Antilles" then response.write (" SELECTED ")%>>
								NetherlandsAntilles
							</option>
							<option <% if COUNTRY="New Caledonia" then response.write (" SELECTED ")%>>
								New Caledonia
							</option>
							<option <% if COUNTRY="New Zealand" then response.write (" SELECTED ")%>>
								New Zealand
							</option>
							<option <% if COUNTRY="Nicaragua" then response.write (" SELECTED ")%>>
								Nicaragua
							</option>
							<option <% if COUNTRY="Niger" then response.write (" SELECTED ")%>>
								Niger
							</option>
							<option <% if COUNTRY="Nigeria" then response.write (" SELECTED ")%>>
								Nigeria
							</option>
							<option <% if COUNTRY="Niue" then response.write (" SELECTED ")%>>
								Niue
							</option>
							<option <% if COUNTRY="Norfolk Island" then response.write (" SELECTED ")%>>
								Norfolk Island
							</option>
							<option <% if COUNTRY="Northern Ireland" then response.write (" SELECTED ")%>>
								Northern Ireland
							</option>
							<option <% if COUNTRY="Northern Mariana Islands" then response.write (" SELECTED ")%>>
								Northern Mariana Islands
							</option>
							<option <% if COUNTRY="Norway" then response.write (" SELECTED ")%>>
								Norway
							</option>
							<option <% if COUNTRY="Oman" then response.write (" SELECTED ")%>>
								Oman
							</option>
							<option <% if COUNTRY="Pakistan" then response.write (" SELECTED ")%>>
								Pakistan
							</option>
							<option <% if COUNTRY="Palau" then response.write (" SELECTED ")%>>
								Palau
							</option>
							<option <% if COUNTRY="Panama" then response.write (" SELECTED ")%>>
								Panama
							</option>
							<option <% if COUNTRY="Papua New Guinea" then response.write (" SELECTED ")%>>
								Papua New Guinea
							</option>
							<option <% if COUNTRY="Paraguay" then response.write (" SELECTED ")%>>
								Paraguay
							</option>
							<option <% if COUNTRY="Peru" then response.write (" SELECTED ")%>>
								Peru
							</option>
							<option <% if COUNTRY="Philippines" then response.write (" SELECTED ")%>>
								Philippines
							</option>
							<option <% if COUNTRY="Pitcairn" then response.write (" SELECTED ")%>>
								Pitcairn
							</option>
							<option <% if COUNTRY="Poland" then response.write (" SELECTED ")%>>
								Poland
							</option>
							<option <% if COUNTRY="Portugal" then response.write (" SELECTED ")%>>
								Portugal
							</option>
							<option <% if COUNTRY="Puerto Rico" then response.write (" SELECTED ")%>>
								Puerto Rico
							</option>
							<option <% if COUNTRY="Qatar" then response.write (" SELECTED ")%>>
								Qatar
							</option>
							<option <% if COUNTRY="Reunion" then response.write (" SELECTED ")%>>
								Reunion
							</option>
							<option <% if COUNTRY="Romania" then response.write (" SELECTED ")%>>
								Romania
							</option>
							<option <% if COUNTRY="Russia" then response.write (" SELECTED ")%>>
								Russia
							</option>
							<option <% if COUNTRY="Russian Federation" then response.write (" SELECTED ")%>>
								Russian Federation
							</option>
							<option <% if COUNTRY="Rwanda" then response.write (" SELECTED ")%>>
								Rwanda
							</option>
							<option <% if COUNTRY="Saint Kitts and Nevis" then response.write (" SELECTED ")%>>
								Saint Kitts and Nevis
							</option>
							<option <% if COUNTRY="Saint Lucia" then response.write (" SELECTED ")%>>
								Saint Lucia
							</option>
							<option <% if COUNTRY="Saint Vincent and the Grenadines" then response.write (" SELECTED ")%>>
								Saint Vincent and the Grenadines
							</option>
							<option <% if COUNTRY="Samoa (Independent)" then response.write (" SELECTED ")%>>
								Samoa (Independent)
							</option>
							<option <% if COUNTRY="San Marino" then response.write (" SELECTED ")%>>
								San Marino
							</option>
							<option <% if COUNTRY="Sao Tome and Principe" then response.write (" SELECTED ")%>>
								Sao Tome and Principe
							</option>
							<option <% if COUNTRY="Saudi Arabia" then response.write (" SELECTED ")%>>
								Saudi Arabia
							</option>
							<option <% if COUNTRY="Scotland" then response.write (" SELECTED ")%>>
								Scotland
							</option>
							<option <% if COUNTRY="Senegal" then response.write (" SELECTED ")%>>
								Senegal
							</option>
							<option <% if COUNTRY="Seychelles" then response.write (" SELECTED ")%>>
								Seychelles
							</option>
							<option <% if COUNTRY="Sierra Leone" then response.write (" SELECTED ")%>>
								SierraLeone
							</option>
							<option <% if COUNTRY="Singapore" then response.write (" SELECTED ")%>>
								Singapore
							</option>
							<option <% if COUNTRY="Slovakia" then response.write (" SELECTED ")%>>
								Slovakia
							</option>
							<option <% if COUNTRY="Slovenia" then response.write (" SELECTED ")%>>
								Slovenia
							</option>
							<option <% if COUNTRY="Solomon Islands" then response.write (" SELECTED ")%>>
								Solomon Islands
							</option>
							<option <% if COUNTRY="Somalia" then response.write (" SELECTED ")%>>
								Somalia
							</option>
							<option <% if COUNTRY="South Africa" then response.write (" SELECTED ")%>>
								South Africa
							</option>
							<option <% if COUNTRY="South Georgia and the South Sandwich Islands" then response.write (" SELECTED ")%>>
								South Georgia and the South Sandwich Islands
							</option>
							<option <% if COUNTRY="South Korea" then response.write (" SELECTED ")%>>
								South Korea
							</option>
							<option <% if COUNTRY="Spain" then response.write (" SELECTED ")%>>
								Spain
							</option>
							<option <% if COUNTRY="Sri Lanka" then response.write (" SELECTED ")%>>
								Sri Lanka
							</option>
							<option <% if COUNTRY="St. Helena" then response.write (" SELECTED ")%>>
								St. Helena
							</option>
							<option <% if COUNTRY="St. Pierre and Miquelon" then response.write (" SELECTED ")%>>
								St. Pierre and Miquelon
							</option>
							<option <% if COUNTRY="Suriname" then response.write (" SELECTED ")%>>
								Suriname
							</option>
							<option <% if COUNTRY="Svalbard and Jan Mayen Islands" then response.write (" SELECTED ")%>>
								Svalbard and Jan Mayen Islands
							</option>
							<option <% if COUNTRY="Swaziland" then response.write (" SELECTED ")%>>
								Swaziland
							</option>
							<option <% if COUNTRY="Sweden" then response.write (" SELECTED ")%>>
								Sweden
							</option>
							<option <% if COUNTRY="Switzerland" then response.write (" SELECTED ")%>>
								Switzerland
							</option>
							<option <% if COUNTRY="Taiwan" then response.write (" SELECTED ")%>>
								Taiwan
							</option>
							<option <% if COUNTRY="Tajikistan" then response.write (" SELECTED ")%>>
								Tajikistan
							</option>
							<option <% if COUNTRY="Tanzania" then response.write (" SELECTED ")%>>
								Tanzania
							</option>
							<option <% if COUNTRY="Thailand" then response.write (" SELECTED ")%>>
								Thailand
							</option>
							<option <% if COUNTRY="Togo" then response.write (" SELECTED ")%>>
								Togo
							</option>
							<option <% if COUNTRY="Tokelau" then response.write (" SELECTED ")%>>
								Tokelau
							</option>
							<option <% if COUNTRY="Tonga" then response.write (" SELECTED ")%>>
								Tonga
							</option>
							<option <% if COUNTRY="Trinidad" then response.write (" SELECTED ")%>>
								Trinidad
							</option>
							<option <% if COUNTRY="Trinidad and Tobago" then response.write (" SELECTED ")%>>
								Trinidad and Tobago
							</option>
							<option <% if COUNTRY="Tunisia" then response.write (" SELECTED ")%>>
								Tunisia
							</option>
							<option <% if COUNTRY="Turkey" then response.write (" SELECTED ")%>>
								Turkey
							</option>
							<option <% if COUNTRY="Turkmenistan" then response.write (" SELECTED ")%>>
								Turkmenistan
							</option>
							<option <% if COUNTRY="Turks and Caicos Islands" then response.write (" SELECTED ")%>>
								Turks and Caicos Islands
							</option>
							<option <% if COUNTRY="Tuvalu" then response.write (" SELECTED ")%>>
								Tuvalu
							</option>
							<option <% if COUNTRY="Uganda" then response.write (" SELECTED ")%>>
								Uganda
							</option>
							<option <% if COUNTRY="Ukraine" then response.write (" SELECTED ")%>>
								Ukraine
							</option>
							<option <% if COUNTRY="United Arab Emirates" then response.write (" SELECTED ")%>>
								United Arab Emirates
							</option>
							<option <% if COUNTRY="United Kingdom" then response.write (" SELECTED ")%>>
								United Kingdom
							</option>
							<option <% if COUNTRY="United States Minor Outlying Islands" then response.write (" SELECTED ")%>>
								United States Minor Outlying Islands
							</option>
							<option <% if COUNTRY="Uruguay" then response.write (" SELECTED ")%>>
								Uruguay
							</option>
							<option <% if COUNTRY="Uzbekistan" then response.write (" SELECTED ")%>>
								Uzbekistan
							</option>
							<option <% if COUNTRY="Vanuatu" then response.write (" SELECTED ")%>>
								Vanuatu
							</option>
							<option <% if COUNTRY="Vatican City State (Holy See)" then response.write (" SELECTED ")%>>
								Vatican City State (Holy See)
							</option>
							<option <% if COUNTRY="Venezuela" then response.write (" SELECTED ")%>>
								Venezuela
							</option>
							<option <% if COUNTRY="Viet Nam" then response.write (" SELECTED ")%>>
								Viet Nam
							</option>
							<option <% if COUNTRY="Virgin Islands (British)" then response.write (" SELECTED ")%>>
								Virgin Islands (British)
							</option>
							<option <% if COUNTRY="Virgin Islands (U.S.)" then response.write (" SELECTED ")%>>
								Virgin Islands (U.S.)
							</option>
							<option <% if COUNTRY="Wales" then response.write (" SELECTED ")%>>
								Wales
							</option>
							<option <% if COUNTRY="Wallis and Futuna Islands" then response.write (" SELECTED ")%>>
								Wallis and Futuna Islands
							</option>
							<option <% if COUNTRY="Western Sahara" then response.write (" SELECTED ")%>>
								Western Sahara
							</option>
							<option <% if COUNTRY="Yemen" then response.write (" SELECTED ")%>>
								Yemen
							</option>
							<option <% if COUNTRY="Zambia" then response.write (" SELECTED ")%>>
								Zambia
							</option>
							<option <% if COUNTRY="Zimbabwe" then response.write (" SELECTED ")%>>
								Zimbabwe
							</option>
						</select>
					</td>
					<td width=100>&nbsp;
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300">&nbsp;
					</td>
				</tr>
<!-- Search by State -->
				<tr>
					<td width=300>
						<div align="right">
							Select State
						</div>
					</td>
					<td width=100>
						<select  name="STATES" onChange="SEARCH.click()">
							<option value="">
								&nbsp;
							</option>
							<option <% if STATES="AL" then response.write (" SELECTED ")%>>
								Alabama
							</option>
							<option <% if STATES="AK" then response.write (" SELECTED ")%>>
								Alaska
							</option>
							<option <% if STATES="AB" then response.write (" SELECTED ")%>>
								Alberta
							</option>
							<option <% if STATES="AZ" then response.write (" SELECTED ")%>>
								Arizona
							</option>
							<option <% if STATES="AR" then response.write (" SELECTED ")%>>
								Arkansas
							</option>
							<option <% if STATES="BC" then response.write (" SELECTED ")%>>
								British Columbia
							</option>
							<option <% if STATES="CA" then response.write (" SELECTED ")%>>
								California
							</option>
							<option <% if STATES="CO" then response.write (" SELECTED ")%>>
								Colorado
							</option>
							<option <% if STATES="CT" then response.write (" SELECTED ")%>>
								Connecticut
							</option>
							<option <% if STATES="DE" then response.write (" SELECTED ")%>>
								Delaware
							</option>
							<option <% if STATES="FL" then response.write (" SELECTED ")%>>
								Florida
							</option>
							<option <% if STATES="GA" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if STATES="HI" then response.write (" SELECTED ")%>>
								Hawaii
							</option>
							<option <% if STATES="ID" then response.write (" SELECTED ")%>>
								Idaho
							</option>
							<option <% if STATES="IL" then response.write (" SELECTED ")%>>
								Illinois
							</option>
							<option <% if STATES="IN" then response.write (" SELECTED ")%>>
								Indiana
							</option>
							<option <% if STATES="IA" then response.write (" SELECTED ")%>>
								Iowa
							</option>
							<option <% if STATES="KS" then response.write (" SELECTED ")%>>
								Kansas
							</option>
							<option <% if STATES="KY" then response.write (" SELECTED ")%>>
								Kentucky
							</option>
							<option <% if STATES="LA" then response.write (" SELECTED ")%>>
								Louisiana
							</option>
							<option <% if STATES="ME" then response.write (" SELECTED ")%>>
								Maine
							</option>
							<option <% if STATES="MB" then response.write (" SELECTED ")%>>
								Manitoba
							</option>
							<option <% if STATES="MD" then response.write (" SELECTED ")%>>
								Maryland
							</option>
							<option <% if STATES="MA" then response.write (" SELECTED ")%>>
								Massachusetts
							</option>
							<option <% if STATES="MI" then response.write (" SELECTED ")%>>
								Michigan
							</option>
							<option <% if STATES="MN" then response.write (" SELECTED ")%>>
								Minnesota
							</option>
							<option <% if STATES="MS" then response.write (" SELECTED ")%>>
								Mississippi
							</option>
							<option <% if STATES="MO" then response.write (" SELECTED ")%>>
								Missouri
							</option>
							<option <% if STATES="MT" then response.write (" SELECTED ")%>>
								Montana
							</option>
							<option <% if STATES="NE" then response.write (" SELECTED ")%>>
								Nebraska
							</option>
							<option <% if STATES="NV" then response.write (" SELECTED ")%>>
								Nevada
							</option>
							<option <% if STATES="NB" then response.write (" SELECTED ")%>>
								New Brunswick
							</option>
							<option <% if STATES="NH" then response.write (" SELECTED ")%>>
								New Hampshire
							</option>
							<option <% if STATES="NJ" then response.write (" SELECTED ")%>>
								New Jersey
							</option>
							<option <% if STATES="NM" then response.write (" SELECTED ")%>>
								New Mexico
							</option>
							<option <% if STATES="NSW" then response.write (" SELECTED ")%>>
								New South Wales
							</option>
							<option <% if STATES="NY" then response.write (" SELECTED ")%>>
								New York
							</option>
							<option <% if STATES="NL" then response.write (" SELECTED ")%>>
								Newfoundland
							</option>
							<option <% if STATES="NC" then response.write (" SELECTED ")%>>
								North Carolina
							</option>
							<option <% if STATES="ND" then response.write (" SELECTED ")%>>
								North Dakota
							</option>
							<option <% if STATES="NT" then response.write (" SELECTED ")%>>
								Northern Territory
							</option>
							<option <% if STATES="NT" then response.write (" SELECTED ")%>>
								Northwest Territories
							</option>
							<option <% if STATES="NS" then response.write (" SELECTED ")%>>
								Nova Scotia
							</option>
							<option <% if STATES="NU" then response.write (" SELECTED ")%>>
								Nunavut
							</option>
							<option <% if STATES="OH" then response.write (" SELECTED ")%>>
								Ohio
							</option>
							<option <% if STATES="OK" then response.write (" SELECTED ")%>>
								Oklahoma
							</option>
							<option <% if STATES="ON" then response.write (" SELECTED ")%>>
								Ontario
							</option>
							<option <% if STATES="OR" then response.write (" SELECTED ")%>>
								Oregon
							</option>
							<option <% if STATES="PA" then response.write (" SELECTED ")%>>
								Pennsylvania
							</option>
							<option <% if STATES="PE" then response.write (" SELECTED ")%>>
								Prince Edward Island
							</option>
							<option <% if STATES="QC" then response.write (" SELECTED ")%>>
								Quebec
							</option>
							<option <% if STATES="QLD" then response.write (" SELECTED ")%>>
								Queensland
							</option>
							<option <% if STATES="RI" then response.write (" SELECTED ")%>>
								Rhode Island
							</option>
							<option <% if STATES="SK" then response.write (" SELECTED ")%>>
								Saskatchewan
							</option>
							<option <% if STATES="SA" then response.write (" SELECTED ")%>>
								South Australia
							</option>
							<option <% if STATES="SC" then response.write (" SELECTED ")%>>
								South Carolina
							</option>
							<option <% if STATES="SD" then response.write (" SELECTED ")%>>
								South Dakota
							</option>
							<option <% if STATES="TAS" then response.write (" SELECTED ")%>>
								Tasmania
							</option>
							<option <% if STATES="TN" then response.write (" SELECTED ")%>>
								Tennessee
							</option>
							<option <% if STATES="TX" then response.write (" SELECTED ")%>>
								Texas
							</option>
							<option <% if STATES="UT" then response.write (" SELECTED ")%>>
								Utah
							</option>
							<option <% if STATES="VT" then response.write (" SELECTED ")%>>
								Vermont
							</option>
							<option <% if STATES="VIC" then response.write (" SELECTED ")%>>
								Victoria
							</option>
							<option <% if STATES="VA" then response.write (" SELECTED ")%>>
								Virginia
							</option>
							<option <% if STATES="WA" then response.write (" SELECTED ")%>>
								Washington
							</option>
							<option <% if STATES="DC" then response.write (" SELECTED ")%>>
								Washington DC
							</option>
							<option <% if STATES="WV" then response.write (" SELECTED ")%>>
								West Virginia
							</option>
							<option <% if STATES="W AUST" then response.write (" SELECTED ")%>>
								Western Australia
							</option>
							<option <% if STATES="WI" then response.write (" SELECTED ")%>>
								Wisconsin
							</option>
							<option <% if STATES="WY" then response.write (" SELECTED ")%>>
								Wyoming
							</option>
							<option <% if STATES="YT" then response.write (" SELECTED ")%>>
								Yukon Territory
							</option>
							<option <% if STATES="Alabama" then response.write (" SELECTED ")%>>
								Alabama
							</option>
							<option <% if STATES="Alaska" then response.write (" SELECTED ")%>>
								Alaska
							</option>
							<option <% if STATES="Alberta" then response.write (" SELECTED ")%>>
								Alberta
							</option>
							<option <% if STATES="Arizona" then response.write (" SELECTED ")%>>
								Arizona
							</option>
							<option <% if STATES="Arkansas" then response.write (" SELECTED ")%>>
								Arkansas
							</option>
							<option <% if STATES="British Columbia" then response.write (" SELECTED ")%>>
								British Columbia
							</option>
							<option <% if STATES="California" then response.write (" SELECTED ")%>>
								California
							</option>
							<option <% if STATES="Colorado" then response.write (" SELECTED ")%>>
								Colorado
							</option>
							<option <% if STATES="Connecticut" then response.write (" SELECTED ")%>>
								Connecticut
							</option>
							<option <% if STATES="Delaware" then response.write (" SELECTED ")%>>
								Delaware
							</option>
							<option <% if STATES="Florida" then response.write (" SELECTED ")%>>
								Florida
							</option>
							<option <% if STATES="Georgia" then response.write (" SELECTED ")%>>
								Georgia
							</option>
							<option <% if STATES="Hawaii" then response.write (" SELECTED ")%>>
								Hawaii
							</option>
							<option <% if STATES="Idaho" then response.write (" SELECTED ")%>>
								Idaho
							</option>
							<option <% if STATES="Illinois" then response.write (" SELECTED ")%>>
								Illinois
							</option>
							<option <% if STATES="Indiana" then response.write (" SELECTED ")%>>
								Indiana
							</option>
							<option <% if STATES="Iowa" then response.write (" SELECTED ")%>>
								Iowa
							</option>
							<option <% if STATES="Kansas" then response.write (" SELECTED ")%>>
								Kansas
							</option>
							<option <% if STATES="Kentucky" then response.write (" SELECTED ")%>>
								Kentucky
							</option>
							<option <% if STATES="Louisiana" then response.write (" SELECTED ")%>>
								Louisiana
							</option>
							<option <% if STATES="Maine" then response.write (" SELECTED ")%>>
								Maine
							</option>
							<option <% if STATES="Manitoba" then response.write (" SELECTED ")%>>
								Manitoba
							</option>
							<option <% if STATES="Maryland" then response.write (" SELECTED ")%>>
								Maryland
							</option>
							<option <% if STATES="Massachusetts" then response.write (" SELECTED ")%>>
								Massachusetts
							</option>
							<option <% if STATES="Michigan" then response.write (" SELECTED ")%>>
								Michigan
							</option>
							<option <% if STATES="Minnesota" then response.write (" SELECTED ")%>>
								Minnesota
							</option>
							<option <% if STATES="Mississippi" then response.write (" SELECTED ")%>>
								Mississippi
							</option>
							<option <% if STATES="Missouri" then response.write (" SELECTED ")%>>
								Missouri
							</option>
							<option <% if STATES="Montana" then response.write (" SELECTED ")%>>
								Montana
							</option>
							<option <% if STATES="Nebraska" then response.write (" SELECTED ")%>>
								Nebraska
							</option>
							<option <% if STATES="Nevada" then response.write (" SELECTED ")%>>
								Nevada
							</option>
							<option <% if STATES="New Brunswick" then response.write (" SELECTED ")%>>
								New Brunswick
							</option>
							<option <% if STATES="New Hampshire" then response.write (" SELECTED ")%>>
								New Hampshire
							</option>
							<option <% if STATES="New Jersey" then response.write (" SELECTED ")%>>
								New Jersey
							</option>
							<option <% if STATES="New Mexico" then response.write (" SELECTED ")%>>
								New Mexico
							</option>
							<option <% if STATES="New South Wales" then response.write (" SELECTED ")%>>
								New South Wales
							</option>
							<option <% if STATES="New York" then response.write (" SELECTED ")%>>
								New York
							</option>
							<option <% if STATES="Newfoundland" then response.write (" SELECTED ")%>>
								Newfoundland
							</option>
							<option <% if STATES="North Carolina" then response.write (" SELECTED ")%>>
								North Carolina
							</option>
							<option <% if STATES="North Dakota" then response.write (" SELECTED ")%>>
								North Dakota
							</option>
							<option <% if STATES="Northern Territory" then response.write (" SELECTED ")%>>
								Northern Territory
							</option>
							<option <% if STATES="Northwest Territories" then response.write (" SELECTED ")%>>
								Northwest Territories
							</option>
							<option <% if STATES="Nova Scotia" then response.write (" SELECTED ")%>>
								Nova Scotia
							</option>
							<option <% if STATES="Nunavut" then response.write (" SELECTED ")%>>
								Nunavut
							</option>
							<option <% if STATES="Ohio" then response.write (" SELECTED ")%>>
								Ohio
							</option>
							<option <% if STATES="Oklahoma" then response.write (" SELECTED ")%>>
								Oklahoma
							</option>
							<option <% if STATES="Ontario" then response.write (" SELECTED ")%>>
								Ontario
							</option>
							<option <% if STATES="Oregon" then response.write (" SELECTED ")%>>
								Oregon
							</option>
							<option <% if STATES="Pennsylvania" then response.write (" SELECTED ")%>>
								Pennsylvania
							</option>
							<option <% if STATES="Prince Edward Island" then response.write (" SELECTED ")%>>
								Prince Edward Island
							</option>
							<option <% if STATES="Quebec" then response.write (" SELECTED ")%>>
								Quebec
							</option>
							<option <% if STATES="Queensland" then response.write (" SELECTED ")%>>
								Queensland
							</option>
							<option <% if STATES="Rhode Island" then response.write (" SELECTED ")%>>
								Rhode Island
							</option>
							<option <% if STATES="Saskatchewan" then response.write (" SELECTED ")%>>
								Saskatchewan
							</option>
							<option <% if STATES="South Australia" then response.write (" SELECTED ")%>>
								South Australia
							</option>
							<option <% if STATES="South Carolina" then response.write (" SELECTED ")%>>
								South Carolina
							</option>
							<option <% if STATES="South Dakota" then response.write (" SELECTED ")%>>
								South Dakota
							</option>
							<option <% if STATES="Tasmania" then response.write (" SELECTED ")%>>
								Tasmania
							</option>
							<option <% if STATES="Tennessee" then response.write (" SELECTED ")%>>
								Tennessee
							</option>
							<option <% if STATES="Texas" then response.write (" SELECTED ")%>>
								Texas
							</option>
							<option <% if STATES="Utah" then response.write (" SELECTED ")%>>
								Utah
							</option>
							<option <% if STATES="Vermont" then response.write (" SELECTED ")%>>
								Vermont
							</option>
							<option <% if STATES="Victoria" then response.write (" SELECTED ")%>>
								Victoria
							</option>
							<option <% if STATES="Virginia" then response.write (" SELECTED ")%>>
								Virginia
							</option>
							<option <% if STATES="Washington" then response.write (" SELECTED ")%>>
								Washington
							</option>
							<option <% if STATES="Washington DC" then response.write (" SELECTED ")%>>
								Washington DC
							</option>
							<option <% if STATES="West Virginia" then response.write (" SELECTED ")%>>
								West Virginia
							</option>
							<option <% if STATES="Western Australia" then response.write (" SELECTED ")%>>
								Western Australia
							</option>
							<option <% if STATES="Wisconsin" then response.write (" SELECTED ")%>>
								Wisconsin
							</option>
							<option <% if STATES="Wyoming" then response.write (" SELECTED ")%>>
								Wyoming
							</option>
							<option <% if STATES="Yukon Territory" then response.write (" SELECTED ")%>>
								Yukon Territory
							</option>
						</select>
					</td>
					<td width=100>&nbsp;
					</td>
				</tr>
				<tr>
					<td width="300">
					</td>
					<td width="300">
					</td>
				</tr>
				<tr>
					<td width="300">&nbsp;
					</td>
					<td width="300" align="left">
						<a href="default.asp">
							Return to Registration
						</a>
					</td>
					<td width="100">&nbsp;
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
