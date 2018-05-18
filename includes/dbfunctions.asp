
<%

Function ManageQuotes(str)
	Dim returnstr
	'If str = "" Then
	'    returnstr = ""
	'Else
	    returnstr = Replace(str, "'", "''")
	'End If
	ManageQuotes = returnstr
End Function

Function ManagePlural(str)
    If InStrRev(LCase(str), "s") <> 0 Then
        str = Mid(str, 1, Len(str) - 1)
    End If
	ManagePlural = str
End Function

%>

