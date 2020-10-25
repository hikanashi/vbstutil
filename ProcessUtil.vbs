Option Explicit

Dim runProc
runProc = IsRunProcesss(WScript.ScriptName)

If runProc = True Then
	Wscript.Echo WScript.ScriptName & "は起動している"
Else
	Wscript.Echo WScript.ScriptName & "は起動してない"
End If

call KillProcesss(WScript.ScriptName)

Function IsRunProcesss(ScriptName)
	Const ProcessName = "wscript.exe"

	IsRunProcesss = False
	'自プロセスは終了対象から除くので自プロセスIDを取得
	Dim ProcessId
	ProcessId = CurrProcessId

	'WMIにて使用する各種オブジェクトを定義・生成する。
	Dim oClassSet
	Dim oClass
	Dim oLocator
	Dim oService

	'ローカルコンピュータに接続する。
	Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set oService = oLocator.ConnectServer
	'クエリー条件を WQL にて指定する。
	Set oClassSet = oService.ExecQuery("Select * From Win32_Process Where Description=""" & ProcessName & """")
	'コレクションを解析する。
	For Each oClass In oClassSet
		Dim lngPos
		lngPos = InStr(oClass.CommandLine, ScriptName)

		If lngPos <> 0 and oClass.ProcessId <> ProcessId Then
			IsRunProcesss = True
		End If
	Next

	'使用した各種オブジェクトを後片付けする。
	Set oClassSet = Nothing
	Set oClass = Nothing
	Set oService = Nothing
	Set oLocator = Nothing
End Function

Function KillProcesss(ScriptName)
	Const ProcessName = "wscript.exe"

	'自プロセスは終了対象から除くので自プロセスIDを取得
	Dim ProcessId
	ProcessId = CurrProcessId

	'WMIにて使用する各種オブジェクトを定義・生成する。
	Dim oClassSet
	Dim oClass
	Dim oLocator
	Dim oService

	'ローカルコンピュータに接続する。
	Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set oService = oLocator.ConnectServer
	'クエリー条件を WQL にて指定する。
	Set oClassSet = oService.ExecQuery("Select * From Win32_Process Where Description=""" & ProcessName & """")
	'コレクションを解析する。
	For Each oClass In oClassSet
		Dim lngPos
		lngPos = InStr(oClass.CommandLine, ScriptName)

		If lngPos <> 0 and oClass.ProcessId <> ProcessId Then
			oClass.Terminate
		End If
	Next

	'使用した各種オブジェクトを後片付けする。
	Set oClassSet = Nothing
	Set oClass = Nothing
	Set oService = Nothing
	Set oLocator = Nothing
End Function

'自プロセスIDを取得する(自プロセスは直接取得できないので子プロセスを作って親プロセスIDを取得する)
Function CurrProcessId
	Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut

	lOut = 0
	Set oShell = CreateObject("WScript.Shell")
	Set oWMI = GetObject(_
		"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	sCmd = "/K " & Left(CreateObject("Scriptlet.TypeLib").Guid, 38)

	oShell.Run "%comspec% " & sCmd, 0
	WScript.Sleep 100 'For healthier skin, get some sleep

	Set oChldPrcs = oWMI.ExecQuery(_
		"Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'",,32)
	For Each oCols In oChldPrcs
		lOut = oCols.ParentProcessId 'get parent
		oCols.Terminate 'process terminated
		Exit For
	Next
	CurrProcessId = lOut
End Function

