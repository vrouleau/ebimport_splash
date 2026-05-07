' Mark Masters athletes: finds athletes with LICENSE ending in '_MA',
' sets BONUSENTRY='T' on their SWIMRESULT rows, then strips the suffix.
'
' Run AFTER importing Lenex into SPLASH, BEFORE generating heats.
' Usage: cscript mark_masters.vbs "C:\path\to\meet.mdb"

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript mark_masters.vbs <path_to_meet.mdb>"
    WScript.Quit 1
End If

Dim mdbPath, conn, rs
Dim marked, cleaned

mdbPath = WScript.Arguments(0)
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mdbPath & ";"

' Find athletes with LICENSE ending in _MA
marked = 0
Set rs = conn.Execute("SELECT ATHLETEID, LICENSE FROM ATHLETE WHERE LICENSE LIKE '%_MA'")
Do While Not rs.EOF
    Dim aid, lic, cleanLic
    aid = CLng(rs("ATHLETEID"))
    lic = rs("LICENSE") & ""
    cleanLic = Left(lic, Len(lic) - 3)

    ' Set BONUSENTRY='T' on all their SWIMRESULT rows
    conn.Execute "UPDATE SWIMRESULT SET BONUSENTRY='T' WHERE ATHLETEID=" & aid

    ' Strip _MA from LICENSE
    conn.Execute "UPDATE ATHLETE SET LICENSE='" & cleanLic & "' WHERE ATHLETEID=" & aid

    marked = marked + 1
    rs.MoveNext
Loop
rs.Close

conn.Close
WScript.Echo "Marked " & marked & " athlete(s) as Masters (BONUSENTRY='T')"
WScript.Echo "LICENSE suffix '_MA' removed."
WScript.Echo ""
WScript.Echo "Now generate heats in SPLASH, run prelims, then run masters_transfer.bat"
