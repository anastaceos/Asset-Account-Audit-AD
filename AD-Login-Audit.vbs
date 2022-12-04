'
' Ensure to update <DOMAIN> with your environments domain name
'
'
' Set Environment Variables 
Set WSHNetwork = WScript.CreateObject("WScript.Network") 
'Set WSHShell = WScript.CreateObject("WScript.Shell") 
'Set WshNetwork = WScript.CreateObject("WScript.Network")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_ComputerSystem")
'Set WSHShell = WScript.CreateObject("WScript.Shell")
'Set objFSO = CreateObject("Scripting.FileSystemObject")

' Determine System Type
For each System in SystemSet
	Select Case System.DomainRole
	  Case 1 Call MemberWSScript()
	  Case Else 'Do Nothing 	'Unknown type
	End Select
	
Exit For
next
wscript.quit

' Member Workstation Script

Function MemberWSScript
			'On Error Resume Next
			Set WshNetwork = WScript.CreateObject("WScript.Network")
			Set objADSysInfo = CreateObject("ADSystemInfo")
			Set WSHShell = CreateObject("WScript.Shell")
			computerName = WshNetwork.ComputerName
			'Initialize global variables
			SiteName = objADSysInfo.SiteName
			ExecuteGlobal "userName = WshNetwork.UserName"
			ExecuteGlobal "userAdsPath = searchAD(""AdsPath"",""user"",userName)"
			ExecuteGlobal "userdisplayName = searchAD(""displayName"",""user"",userName)"
			ExecuteGlobal "Set userObj = GetObject(userAdsPath)"
			Call UpdateLastLogon()
			'wscript.sleep 1000
			Call Adduser2computerdescriptionAD()
End Function

' AD User Object Description Update
Function UpdateLastLogon()
' Popup message
	Set objShell = Wscript.CreateObject("Wscript.Shell")
	objShell.Popup "Registering this computer with <DOMAIN>",5,"<DOMAIN> Active Directory Update",64
' Updating AD info for user
	Set objSysInfo = CreateObject("ADSystemInfo")
	Set WSHShell = CreateObject("WScript.Shell")
	Set WSHSysEnv = WSHShell.Environment("PROCESS")
	Set g_objADObject = GetObject("LDAP://" & objSysInfo.UserName)
	g_objADObject.Put "info", "Last logged on at: " & Now() & " on: " & WSHSysEnv("COMPUTERNAME") 
    g_objADObject.SetInfo
End Function

' AD Computer Object Description Update
Function Adduser2computerdescriptionAD()
' Updating AD info for computer
   Dim myIPAddress : myIPAddress = ""
   Dim strFullName
   Dim objWMIService : Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
   Dim colAdapters : Set colAdapters = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
   Dim objAdapter
      For Each objAdapter in colAdapters
        If Not IsNull(objAdapter.IPAddress) Then myIPAddress = trim(objAdapter.IPAddress(0))
      exit for
     Next
 		
   Set objSysInfo = CreateObject("ADSystemInfo")
   Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
   Set objComputer = GetObject("LDAP://" & objSysInfo.ComputerName)
     strMessage = objUser.CN & " - " & objUser.DisplayName & " logged in @ IP:" & myIPAddress & "   " & Now
     strFullName = objUser.DisplayName 
	  'wscript.echo strMessage
      'objUser.Description = strMessage
      'objUser.SetInfo
   objComputer.Description = strMessage
   'objComputer.employeeid = objUser.CN
   objComputer.SetInfo

   Set objShell = Wscript.CreateObject("Wscript.Shell")
objShell.Popup "Update Successful",5,"<DOMAIN> Active Directory Update",64
				 
End Function

' Sub-Routines for  Workstation Script

Function searchAD(attrib,category,cnName)
				' Returns the object attributes you specify
				' for the object named 'cnName' of type 'category'

		' Create the connection and command object.
	Set oConnection1 = CreateObject("ADODB.Connection")
	Set oCommand1 = CreateObject("ADODB.Command")
		' Open the connection.
	oConnection1.Provider = "ADsDSOObject"  ' This is the ADSI OLE-DB provider name
	oConnection1.Open "Active Directory Provider"
		' Create a command object for this connection.
	Set oCommand1.ActiveConnection = oConnection1
		' Compose a search string.
	oCommand1.CommandText = "select "&attrib&" from 'LDAP://<DOMAIN>.com/DC=<DOMAIN>,DC=com' WHERE objectCategory='"&category&"' AND cn='"&cnName&"'"

		' Execute the query.
	Set rs = oCommand1.Execute
		' Navigate the record set
	While Not rs.EOF
		For i = 0 To rs.Fields.Count - 1
			tmp=tmp+rs.Fields(i).Value+" "
				'MsgBox rs.Fields(i).Name & " = " & rs.Fields(i).Value
		Next 
		rs.MoveNext
	Wend
	searchAD = tmp
End Function

' End
