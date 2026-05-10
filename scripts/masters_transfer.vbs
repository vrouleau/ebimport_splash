' Masters Time Transfer for SPLASH Meet Manager
' Copies Masters athletes from prelim events to Masters finals,
' creates heats, and deletes the prelim rows.
'
' Identifies Masters by HANDICAP exception='X' (set during Lenex import).
'
' Usage: cscript masters_transfer.vbs "C:\path\to\meet.mdb"
' Run AFTER prelim heats have been swum (SWIMTIME populated).

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript masters_transfer.vbs <path_to_meet.mdb>"
    WScript.Quit 1
End If

Dim mdbPath, conn, rs, rs2, sql
Dim nextUID, laneMin, laneMax, lanesPerHeat, ageDate
Dim totalTransferred, totalHeats, totalDeleted
Dim mvData, adPos, adStr
Dim dictFinals, fk
Dim pEids(100), pStyles(100), pGenders(100), pEnums(100), pCount, pi
Dim lk, fi, finalEid, finalEnum
Dim fagIds(50), fagMins(50), fagMaxs(50), fagCount
Dim maxHeat
Dim srIds(500), athIds(500), swimtimes(500), entrytimes(500)
Dim reactions(500), statuses(500), birthdates(500), srCount
Dim heatNum, laneNum, currentHeatId, eventCount, si
Dim athleteAge, bdVal, targetAGID, ki, stVal
Dim useBonus, rsMA, maCount, maAid
Dim shouldInclude, tmpAge

mdbPath = WScript.Arguments(0)
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mdbPath & ";"

' Get next UID
Set rs = conn.Execute("SELECT LASTUID FROM BSUIDTABLE WHERE NAME='BS_GLOBAL_UID'")
nextUID = CLng(rs("LASTUID")) + 1
rs.Close

' Get lane config
Set rs = conn.Execute("SELECT TOP 1 LANEMIN, LANEMAX FROM SWIMSESSION")
laneMin = CInt(rs("LANEMIN"))
laneMax = CInt(rs("LANEMAX"))
rs.Close
lanesPerHeat = laneMax - laneMin + 1

' Get AGEDATE
Set rs = conn.Execute("SELECT DATA FROM BSGLOBAL WHERE NAME='MEETVALUES'")
mvData = rs("DATA")
rs.Close
adPos = InStr(mvData, "AGEDATE=D;")
If adPos > 0 Then
    adStr = Mid(mvData, adPos + 10, 8)
    ageDate = DateSerial(CInt(Left(adStr,4)), CInt(Mid(adStr,5,2)), CInt(Mid(adStr,7,2)))
Else
    ageDate = #12/31/2026#
End If
WScript.Echo "Age date: " & ageDate

' Detect Masters athletes via HANDICAP exception='X'
' Mark their SWIMRESULT rows with BONUSENTRY='T' for transfer
maCount = 0
Set rsMA = conn.Execute("SELECT ATHLETEID FROM ATHLETE WHERE EXCEPTION='X'")
Do While Not rsMA.EOF
    maAid = CLng(rsMA("ATHLETEID"))
    conn.Execute "UPDATE SWIMRESULT SET BONUSENTRY='T' WHERE ATHLETEID=" & maAid
    maCount = maCount + 1
    rsMA.MoveNext
Loop
rsMA.Close
If maCount > 0 Then
    WScript.Echo "Marked " & maCount & " Masters athletes (exception=X)"
End If

' Check BONUSENTRY count
Set rs = conn.Execute("SELECT COUNT(*) AS C FROM SWIMRESULT WHERE BONUSENTRY='T'")
If CInt(rs("C")) > 0 Then
    useBonus = True
    WScript.Echo "Mode: BONUSENTRY (" & CInt(rs("C")) & " entries)"
Else
    useBonus = False
    WScript.Echo "Mode: age-based fallback (no Masters markers found)"
End If
rs.Close

' Build finals lookup
Set dictFinals = CreateObject("Scripting.Dictionary")
Set rs = conn.Execute("SELECT SWIMEVENTID, SWIMSTYLEID, GENDER, EVENTNUMBER FROM SWIMEVENT WHERE ROUND = 1 AND MASTERS = 'T'")
Do While Not rs.EOF
    fk = CLng(rs("SWIMSTYLEID")) & "_" & CInt(rs("GENDER"))
    If Not dictFinals.Exists(fk) Then
        dictFinals.Add fk, Array(CLng(rs("SWIMEVENTID")), CInt(rs("EVENTNUMBER")))
    End If
    rs.MoveNext
Loop
rs.Close

' Collect prelim events
pCount = 0
Set rs = conn.Execute("SELECT SWIMEVENTID, SWIMSTYLEID, GENDER, EVENTNUMBER FROM SWIMEVENT WHERE ROUND = 2 AND MASTERS = 'F'")
Do While Not rs.EOF
    pEids(pCount) = CLng(rs("SWIMEVENTID"))
    pStyles(pCount) = CLng(rs("SWIMSTYLEID"))
    pGenders(pCount) = CInt(rs("GENDER"))
    pEnums(pCount) = CInt(rs("EVENTNUMBER"))
    pCount = pCount + 1
    rs.MoveNext
Loop
rs.Close

WScript.Echo "Prelim events: " & pCount & ", Finals: " & dictFinals.Count

totalTransferred = 0 : totalHeats = 0 : totalDeleted = 0

For pi = 0 To pCount - 1
    lk = pStyles(pi) & "_" & pGenders(pi)
    If dictFinals.Exists(lk) Then
        fi = dictFinals(lk)
        finalEid = fi(0)
        finalEnum = fi(1)

        ' Get final event agegroups
        fagCount = 0
        Set rs = conn.Execute("SELECT AGEGROUPID, AGEMIN, AGEMAX FROM AGEGROUP WHERE SWIMEVENTID=" & finalEid & " AND AGEMIN >= 25 AND AGEMIN < 100")
        Do While Not rs.EOF
            fagIds(fagCount) = CLng(rs("AGEGROUPID"))
            fagMins(fagCount) = CInt(rs("AGEMIN"))
            If IsNull(rs("AGEMAX")) Or CInt(rs("AGEMAX")) < 0 Then
                fagMaxs(fagCount) = 999
            Else
                fagMaxs(fagCount) = CInt(rs("AGEMAX"))
            End If
            fagCount = fagCount + 1
            rs.MoveNext
        Loop
        rs.Close

        ' Get max heat
        Set rs = conn.Execute("SELECT MAX(HEATNUMBER) AS MH FROM HEAT WHERE SWIMEVENTID=" & finalEid)
        If IsNull(rs("MH")) Then maxHeat = 0 Else maxHeat = CInt(rs("MH"))
        rs.Close

        ' Get Masters athletes with results
        srCount = 0
        If useBonus Then
            sql = "SELECT sr.SWIMRESULTID, sr.ATHLETEID, sr.SWIMTIME, sr.ENTRYTIME, sr.REACTIONTIME, sr.RESULTSTATUS, ath.BIRTHDATE FROM SWIMRESULT sr INNER JOIN ATHLETE ath ON ath.ATHLETEID = sr.ATHLETEID WHERE sr.SWIMEVENTID = " & pEids(pi) & " AND sr.BONUSENTRY = 'T'"
        Else
            sql = "SELECT sr.SWIMRESULTID, sr.ATHLETEID, sr.SWIMTIME, sr.ENTRYTIME, sr.REACTIONTIME, sr.RESULTSTATUS, ath.BIRTHDATE FROM SWIMRESULT sr INNER JOIN ATHLETE ath ON ath.ATHLETEID = sr.ATHLETEID WHERE sr.SWIMEVENTID = " & pEids(pi) & " AND ath.BIRTHDATE IS NOT NULL"
        End If
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
            stVal = rs("SWIMTIME")
            If Not IsNull(stVal) Then
                If CLng(stVal) > 0 Then
                    shouldInclude = False
                    If useBonus Then
                        shouldInclude = True
                    Else
                        ' Age-based: include if age >= 25
                        If Not IsNull(rs("BIRTHDATE")) Then
                            tmpAge = Year(ageDate) - Year(CDate(rs("BIRTHDATE")))
                            If DateSerial(Year(ageDate), Month(CDate(rs("BIRTHDATE"))), Day(CDate(rs("BIRTHDATE")))) > ageDate Then
                                tmpAge = tmpAge - 1
                            End If
                            If tmpAge >= 25 Then shouldInclude = True
                        End If
                    End If
                    If shouldInclude Then
                        srIds(srCount) = CLng(rs("SWIMRESULTID"))
                        athIds(srCount) = CLng(rs("ATHLETEID"))
                        swimtimes(srCount) = CLng(stVal)
                        If IsNull(rs("ENTRYTIME")) Then
                            entrytimes(srCount) = CLng(stVal)
                        Else
                            entrytimes(srCount) = CLng(rs("ENTRYTIME"))
                        End If
                        If IsNull(rs("REACTIONTIME")) Then reactions(srCount) = -32768 Else reactions(srCount) = CLng(rs("REACTIONTIME"))
                        If IsNull(rs("RESULTSTATUS")) Then statuses(srCount) = 0 Else statuses(srCount) = CInt(rs("RESULTSTATUS"))
                        birthdates(srCount) = rs("BIRTHDATE")
                        srCount = srCount + 1
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close

        If srCount > 0 Then
            heatNum = maxHeat
            laneNum = laneMax + 1
            currentHeatId = -1
            eventCount = 0

            For si = 0 To srCount - 1
                If Not IsNull(birthdates(si)) And birthdates(si) <> "" Then
                    bdVal = CDate(birthdates(si))
                    athleteAge = Year(ageDate) - Year(bdVal)
                    If DateSerial(Year(ageDate), Month(bdVal), Day(bdVal)) > ageDate Then
                        athleteAge = athleteAge - 1
                    End If

                    targetAGID = -1
                    For ki = 0 To fagCount - 1
                        If athleteAge >= fagMins(ki) And athleteAge <= fagMaxs(ki) Then
                            targetAGID = fagIds(ki)
                            Exit For
                        End If
                    Next

                    If targetAGID > -1 Then
                        If laneNum > laneMax Then
                            heatNum = heatNum + 1
                            laneNum = laneMin
                            conn.Execute "INSERT INTO HEAT (HEATID, SWIMEVENTID, HEATNUMBER) VALUES (" & nextUID & ", " & finalEid & ", " & heatNum & ")"
                            currentHeatId = nextUID
                            nextUID = nextUID + 1
                            totalHeats = totalHeats + 1
                        End If

                        conn.Execute "INSERT INTO SWIMRESULT (SWIMRESULTID, ATHLETEID, SWIMEVENTID, AGEGROUPID, HEATID, LANE, SWIMTIME, ENTRYTIME, ENTRYCOURSE, REACTIONTIME, RESULTSTATUS) VALUES (" & nextUID & ", " & athIds(si) & ", " & finalEid & ", " & targetAGID & ", " & currentHeatId & ", " & laneNum & ", " & swimtimes(si) & ", " & entrytimes(si) & ", 0, " & reactions(si) & ", " & statuses(si) & ")"
                        nextUID = nextUID + 1
                        laneNum = laneNum + 1
                        totalTransferred = totalTransferred + 1

                        conn.Execute "DELETE FROM SWIMRESULT WHERE SWIMRESULTID = " & srIds(si)
                        totalDeleted = totalDeleted + 1
                        eventCount = eventCount + 1
                    End If
                End If
            Next

            If eventCount > 0 Then
                WScript.Echo "  prelim #" & pEnums(pi) & " -> final #" & finalEnum & ": " & eventCount
            End If
        End If
    End If
Next

conn.Execute "UPDATE BSUIDTABLE SET LASTUID = " & (nextUID - 1) & " WHERE NAME='BS_GLOBAL_UID'"
conn.Close

WScript.Echo ""
WScript.Echo "Done: " & totalTransferred & " transferred, " & totalHeats & " heats, " & totalDeleted & " deleted"
