
<%
dim strCon
dim C_BOOTH_COST
dim C_EXHIBITOR_BADGE_COST

'database connection string
'use when connectiong from outside the db server
strCon = "Provider=SQLOLEDB; Data Source=sql6.media3.net; Initial Catalog=midwestclinic; User Id=midwestclinic; Password=DW2148ahR;"
'use when connecting from the db server itself
'strCon = "Provider=SQLOLEDB; Data Source=192.168.10.1; Initial Catalog=midwestclinic; User Id=sa; Password=sql;"

'for db on server1, local odbc connection
'strCon = "Provider=SQLOLEDB; Data Source=(local); Initial Catalog=midwestclinic; Integrated Security=SSPI"
'strCon = "Provider=SQLOLEDB; Data Source=192.168.10.1,1433; Initial Catalog=midwestclinic; Integrated Security=SSPI"


'booth cost 2007
C_BOOTH_COST = 900
'exhibitor badge cost 2008
C_EXHIBITOR_BADGE_COST = 90
%>
