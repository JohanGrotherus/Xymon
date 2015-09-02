' Active Directory replication check for BBWin/Xymon
'
' By Johan Grotherus <johan.grotherus@gmail.com>
' Version 0.1 2015-08-24

On Error Resume Next

Dim objShell
Dim typelib,ServerName,boolValid
Dim strLayout,strSize,strLine,intStart,strFail,strCmd

' Xymon variables
strTestName = "replication"
strAlarmState = "green"
strOutput   = ""

Set oShell = WScript.CreateObject("WScript.Shell")
' Check to see If used with BBWin or Big Brother client and set extPath.
extPath = oShell.RegRead("HKLM\SOFTWARE\BBWin\tmpPath")
If extPath = "" Then
	extPath = oShell.RegRead("HKLM\SOFTWARE\Quest Software\BigBrother\bbnt\ExternalPath\")
End If

strOutput = strOutput & vbcrlf & "<h2>Active Directory replication status:</h2>" & vbcrlf 

Set ObjSysInfo = CreateObject("ADSystemInfo")
strOutput = strOutput & vbcrlf & "Computer name: " & objSysInfo.ComputerName
strOutput = strOutput & vbcrlf & "Site name: " & objSysInfo.SiteName
strOutput = strOutput & vbcrlf & "Domain shortname: " & objSysInfo.DomainShortName
strOutput = strOutput & vbcrlf & "Domain DNS name: " & objSysInfo.DomainDNSName
strOutput = strOutput & vbcrlf & "Forest DNS name: " & objSysInfo.ForestDNSName
strOutput = strOutput & vbcrlf & "PDC Role owner: " & objSysInfo.PDCRoleOwner
strOutput = strOutput & vbcrlf & "Schema Role owner: " & objSysInfo.SchemaRoleOwner
strOutput = strOutput & vbcrlf & "Domain is in native mode: " & objSysInfo.IsNativeMode
strOutput = strOutput & vbcrlf

' Active Directory WMI queries
strComputer = "."

SET objWMIService = GETOBJECT("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & _
         strComputer & "\root\MicrosoftActiveDirectory")
         
SET colReplicationOperations = objWMIService.ExecQuery _
     ("select * from MSAD_ReplNeighbor")
     
FOR EACH objReplicationJob in colReplicationOperations
     strOutput = strOutput & vbcrlf & "Domain: " & objReplicationJob.Domain
     strOutput = strOutput & vbcrlf & "Naming context DN: " & objReplicationJob.NamingContextDN
     strOutput = strOutput & vbcrlf & "Source DSA DN: " & objReplicationJob.SourceDsaDN
     strOutput = strOutput & vbcrlf & "Last sync result: " & objReplicationJob.LastSyncResult
     strOutput = strOutput & vbcrlf & "Number of consecutive synchronization failues: " & _
             objReplicationJob.NumConsecutiveSyncFailures
     strOutput = strOutput & vbcrlf
NEXT

IF objReplicationJob.NumConsecutiveSyncFailures = "0" THEN
    strOutput = strOutput & vbcrlf & "<h3>Active Directory replication is OK on this domain controller</h3>"
ELSE
    strAlarmState = "red"
    strOutput = strOutput & vbcrlf & "<h3>Active Directory replication errors on this domain controller</h3>"
END IF

' Write the file for BBWin
WriteFile extPath, strTestName, strAlarmState, strOutput

' This SUB is used for outputting the file to the external's directory in bbwin
SUB WriteFile(strExtPath, strTestName, strAlarmState, strOutput)
    Set fso = CreateObject("Scripting.FileSystemObject")	
    strOutput = strAlarmState & " " & Date & " " & LCase(Time) & " " & LCase(ServerName) & vbcrlf & strOutput & vbcrlf
    Set f = fso.OpenTextFile(strExtPath & "\" & strTestName , 8 , TRUE)
    f.Write strOutput
    f.Close
    Set fso = Nothing
END Sub


