<% @Language=VBScript %>
<% Server.ScriptTimeout = 999 %>

<!--#include virtual="registration/includes/dbfunctions.asp" -->
<!--#include virtual="registration/includes/dbglobals.asp" -->
<!--#include virtual="registration/includes/adovbs.inc"-->
<%
	dim sCOMPANY_LIST
	dim lngRecs
	dim strSQL
	dim conDB
	dim rsExhibitor
	dim rsBooths
	dim rsBadgeAmtPaid
	dim rsBadgeInfo
	dim rsNumOfBooths
	dim rsBoothAmtPaid
	dim rsTickets
	dim sBoothPayments
	dim sBadgePayments
	dim sCompany
	dim sCity
	dim sState
	dim sExhibID
	dim sBoothNums
	dim sBoothLocs
	dim sBadgeID
	dim iNum
	dim iNumOfBooths
	dim sBadgeTable
	dim sBadgeNames
	dim sBadgeComps
	dim sBadgeTickets
	dim sBadgePrograms
	dim arCompanyField	
	
	COUNTRY = "United States"
	
	Session("EXHIBITOR_ID") = ""
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<title>The Registration - New Exhibitor</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script LANGUAGE="JavaScript"> 
			function IsItFive()
			{
// when the onKeyUp event is fired for the zip code entry,
// look for 5 numbers and then parse the xmldoc
				if (exhibitor_new.ZIP.value.length == 5)
				{
					searchByZipcode(exhibitor_new.ZIP.value);
				}
			}
			
			function searchByZipcode(zcode)
			{
				var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
				var nodes;
				var allZipcodes;
				var allCities;
				var allStates;
				var vNumber;
				var vCity;
				var vState;
			  
				xmlDoc.async="false";
				
				xmlDoc.load("zipxml.xml");
				
				allZipcodes = xmlDoc.getElementsByTagName("zipcode");
				
				for (var i = 0; i < allZipcodes.length; i++)
				{
					vNumber = allZipcodes.item(i).firstChild.nodeValue;
					
					if (vNumber == zcode)
					{
						allCities = xmlDoc.getElementsByTagName("city");
						
						vCity = allCities.item(i).firstChild.nodeValue;
						
						exhibitor_new.CITY.value = vCity;
						
						allStates = xmlDoc.getElementsByTagName("state");
						
						vState = allStates.item(i).firstChild.nodeValue;
						
						for (var j = 0; j < exhibitor_new.STATES.options.length; j++)
						{
							if (exhibitor_new.STATES.options[j].text == vState)
							{
								exhibitor_new.STATES.selectedIndex = j;
								
								break;
							}
						}
						return;
					}
				}
			}
			
			function validateForm()
			{
				if (exhibitor_new.COMPANY_NAME.value.length == 0)
				{
					alert("Please enter the Company Name.");
					
					return false;
				}
				else if (exhibitor_new.CONTACT_NAME.value.length == 0)
				{
					alert("Please enter the Contact Name.");
					
					return false;
				}
				else if (exhibitor_new.EMAIL.value.length == 0)
				{
					alert("Please enter the Email.");
					
					return false;
				}
				else if (exhibitor_new.ZIP.value.length == 0)
				{
					alert("Please enter the Zip Code.");
					
					return false;
				}
				else if (exhibitor_new.ADDRESS.value.length == 0)
				{
					alert("Please enter the Address.");
					
					return false;
				}
				else if (exhibitor_new.CITY.value.length == 0)
				{
					alert("Please enter the City.");
					
					return false;
				}
				else if (exhibitor_new.PHONE.value.length == 0)
				{
					alert("Please enter the Phone.");
					
					return false;
				}
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
	<body>
<!-- BEGIN CONTENT -->
		<form name="exhibitor_new" method=post onSubmit="return validateForm()" action="save_exhibitor_new.asp">
			<table width="100%" border=0 cellspacing="0">
				<tr bgcolor="#AAD5FF">
					<td colspan=2>
						<strong>
							Exhibitor Information
						</strong>
					</td>
					<td width=10%>
					</td>
					<td width=10%>
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Company Name
						</div>
					</td>
					<td width=35%>
						<input type="text" value="<%=COMPANY_NAME%>" name="COMPANY_NAME" size=50 maxlength=50>
					</td>
					<td width=15%>
						<div align="right">
							Contact Name
						</div>
					</td>
					<td width=35%>
						<input type="text" value="<%=CONTACT_NAME%>" name="CONTACT_NAME" size=50 maxlength=50>
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Address
						</div>
					</td>
					<td width=35%>
						<input type=text value="<%=ADDRESS%>" name="ADDRESS" size=50 maxlength=50>
					</td>
					<td width=15%>&nbsp;
					</td>
					<td width=35%>&nbsp;
					</td>
				</tr>
				<tr>
					<td width=15% height="26">
						<div align="right">
							City
						</div>
					</td>
					<td width=35%>
						<input value="<%=CITY%>" type=text name="CITY" size=50 maxlength=50>
					</td>
					<td width=15%>
						<div align="right">
							State/Province/Region
						</div>
					</td>
					<td width=35%>
						<select value="Select A State" name="STATES">
							<option >
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
							<option <% if STATES="TE" then response.write (" SELECTED ")%>>
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
						</select>
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Zip/Postal Code
						</div>
					</td>
					<td width=35%>
						<input type="text" value="<%=ZIP%>" name="ZIP" size=15 >
					</td>
					<td width=15%>
						<div align="right">
							Phone
						</div>
					</td>
					<td width=35%>
						<input type=text value="<%=PHONE%>" name="PHONE" size=50 maxlength=50>
					</td>
				</tr>
				<tr>
					<td width=15%>
						<div align="right">
							Country
						</div>
					</td>
					<td width=35%>
						<select name="COUNTRY" >
							<option value="<%=COUNTRY%>">
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
								Trinidadand Tobago
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
							<option <% if COUNTRY="United States" then response.write (" SELECTED ")%>>
								United States
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
					<td width=15%>
						<div align="right">
							Email
						</div>
					</td>
					<td width=35%>
						<input type="text" value="<%=EMAIL%>" name="EMAIL" size=50 maxlength=50>
					</td>
				</tr>
			</table>
			<table border=0 width=100%>
				<tr>
					<td width=100%>
					</td>
				</tr>
				<tr>
					<td colspan=2>
						<div align="center">
							<input type="submit" name="Submit" value="Add Booths">
							<input type="reset" name="Reset" value="Reset">
						</div>
					</td>
				</tr>
				<tr>
					<td width=100%>
						<div align="center">
							<a href="select_exhibitor.asp">
								Return to Select Exhibitors
							</a>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
