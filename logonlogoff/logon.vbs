Option Explicit
Set objShell=CreateObject("WScript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\xampp\htdocs\Projetos\Sistemas\LogoonLogoff\log\logon.LOG", ForAppending, True)
Dim objWMIService, objItem, objService, objShell, connection, recordSet, connectionString
Dim strComputerName, strService, colServiceList, FileSysObj, WSHShell, objFSO, objTextFile, strDate, strTime, strUserName, strSessionName, strLogonServer, DateInfo
'strDate = wshShell.ExpandEnvironmentStrings( "%Date%" )
'strTime = wshShell.ExpandEnvironmentStrings( "%time%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%ComputerName%" )
strUserName = wshShell.ExpandEnvironmentStrings( "%Username%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%sessionname%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%logonserver%" )
DateInfo = DateInfo & Now & VbCrLf
'DateInfo = DateInfo & Date & VbCrLf
'DateInfo = DateInfo & Time & VbCrLf

Const OverWriteFiles = True
Const ForAppending = 8

objTextFile.WriteLine("Logon Realizado:" & ";" & strUserName & ";" & strComputerName & ";" & DateInfo)


rem dim connection, recordSet, connectionString
set connection = CreateObject("ADODB.Connection")
set recordSet = CreateObject("ADODB.Recordset")

' To list all MySQL ODBC drivers installed on your machine,
' Use PowerShell Get-OdbcDriver -Name "MySQL*"
connectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;" & _
    "Database=logonlogoff;User=root;Password=;"

connection.Open connectionString
recordSet.Open "INSERT INTO `logon` (`codigo`, `tipo`, `usuario`, `Computador`, `datahora`) VALUES (NULL, 'logon', '" & strUserName & "','" & strComputerName & "', '" & DateInfo & "');", connection
