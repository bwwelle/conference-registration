
<%
dim strCon
dim C_BOOTH_COST
dim C_EXHIBITOR_BADGE_COST

'database connection string
'use when connectiong from outside the db server
'strCon = "Provider=SQLOLEDB; Data Source=; Initial Catalog=midwestclinic; User Id=; Password=;"
'use when connecting from the db server itself
strCon = "Provider=SQLOLEDB; Data Source=; Initial Catalog=; User Id=; Password=;"

'for db on server1, local odbc connection
'strCon = "Provider=SQLOLEDB; Data Source=; Initial Catalog=; Integrated Security=SSPI"
'strCon = "Provider=SQLOLEDB; Data Source=; Initial Catalog=; Integrated Security=SSPI"


'booth cost 2007
C_BOOTH_COST = 900
'exhibitor badge cost 2008
C_EXHIBITOR_BADGE_COST = 90
%>
