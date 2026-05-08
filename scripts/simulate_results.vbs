' Simulate Results: write random SWIMTIME into all SWIMRESULT and RELAY rows
' For each entry: SWIMTIME = ENTRYTIME +/- 5%, or random if NT
' 5% get DQ (RESULTSTATUS=2)
'
' Usage: cscript simulate_results.vbs "C:\path\to\meet.mdb"

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript simulate_results.vbs <path_to_meet.mdb>"
    WScript.Quit 1
End If

Dim mdbPath, conn, rs
Dim i, swimTime, baseTime, variation, status

mdbPath = WScript.Arguments(0)
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mdbPath & ";"

Randomize

' --- Individual results ---
Dim srIds(5000), entryTimes(5000), srCount
srCount = 0
Set rs = conn.Execute("SELECT SWIMRESULTID, ENTRYTIME, SWIMTIME FROM SWIMRESULT")
Do While Not rs.EOF
    If IsNull(rs("SWIMTIME")) Or CLng(rs("SWIMTIME")) = 0 Then
        srIds(srCount) = CLng(rs("SWIMRESULTID"))
        If IsNull(rs("ENTRYTIME")) Then
            entryTimes(srCount) = 0
        Else
            entryTimes(srCount) = CLng(rs("ENTRYTIME"))
        End If
        srCount = srCount + 1
    End If
    rs.MoveNext
Loop
rs.Close

Dim totalResults, totalDQ
totalResults = 0 : totalDQ = 0

For i = 0 To srCount - 1
    baseTime = entryTimes(i)
    If baseTime > 0 And baseTime < 2147483647 Then
        variation = baseTime * 0.05
        swimTime = CLng(baseTime + (Rnd * 2 - 1) * variation)
        If swimTime < 1000 Then swimTime = 1000
    Else
        swimTime = CLng(30000 + Rnd * 150000)
    End If
    If Rnd < 0.05 Then
        status = 2
        totalDQ = totalDQ + 1
    Else
        status = 0
    End If
    conn.Execute "UPDATE SWIMRESULT SET SWIMTIME=" & swimTime & ", RESULTSTATUS=" & status & " WHERE SWIMRESULTID=" & srIds(i)
    totalResults = totalResults + 1
Next

WScript.Echo "  " & totalResults & " individual results (" & totalDQ & " DQ)"

' --- Relay results ---
Dim rlIds(2000), rlEntryTimes(2000), rlCount
rlCount = 0
Set rs = conn.Execute("SELECT RELAYID, ENTRYTIME, SWIMTIME FROM RELAY")
Do While Not rs.EOF
    If IsNull(rs("SWIMTIME")) Or CLng(rs("SWIMTIME")) = 0 Then
        rlIds(rlCount) = CLng(rs("RELAYID"))
        If IsNull(rs("ENTRYTIME")) Then
            rlEntryTimes(rlCount) = 0
        Else
            rlEntryTimes(rlCount) = CLng(rs("ENTRYTIME"))
        End If
        rlCount = rlCount + 1
    End If
    rs.MoveNext
Loop
rs.Close

Dim totalRelays, totalRelayDQ
totalRelays = 0 : totalRelayDQ = 0

For i = 0 To rlCount - 1
    baseTime = rlEntryTimes(i)
    If baseTime > 0 And baseTime < 2147483647 Then
        variation = baseTime * 0.05
        swimTime = CLng(baseTime + (Rnd * 2 - 1) * variation)
        If swimTime < 1000 Then swimTime = 1000
    Else
        swimTime = CLng(120000 + Rnd * 180000)
    End If
    If Rnd < 0.05 Then
        status = 2
        totalRelayDQ = totalRelayDQ + 1
    Else
        status = 0
    End If
    conn.Execute "UPDATE RELAY SET SWIMTIME=" & swimTime & ", RESULTSTATUS=" & status & " WHERE RELAYID=" & rlIds(i)
    totalRelays = totalRelays + 1
Next

conn.Close
WScript.Echo "  " & totalRelays & " relay results (" & totalRelayDQ & " DQ)"
WScript.Echo "Done."
