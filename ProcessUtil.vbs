Option Explicit

Dim runProc
runProc = IsRunProcesss(WScript.ScriptName)

If runProc = True Then
	Wscript.Echo WScript.ScriptName & "�͋N�����Ă���"
Else
	Wscript.Echo WScript.ScriptName & "�͋N�����ĂȂ�"
End If

call KillProcesss(WScript.ScriptName)

Function IsRunProcesss(ScriptName)
	Const ProcessName = "wscript.exe"

	IsRunProcesss = False
	'���v���Z�X�͏I���Ώۂ��珜���̂Ŏ��v���Z�XID���擾
	Dim ProcessId
	ProcessId = CurrProcessId

	'WMI�ɂĎg�p����e��I�u�W�F�N�g���`�E��������B
	Dim oClassSet
	Dim oClass
	Dim oLocator
	Dim oService

	'���[�J���R���s���[�^�ɐڑ�����B
	Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set oService = oLocator.ConnectServer
	'�N�G���[������ WQL �ɂĎw�肷��B
	Set oClassSet = oService.ExecQuery("Select * From Win32_Process Where Description=""" & ProcessName & """")
	'�R���N�V��������͂���B
	For Each oClass In oClassSet
		Dim lngPos
		lngPos = InStr(oClass.CommandLine, ScriptName)

		If lngPos <> 0 and oClass.ProcessId <> ProcessId Then
			IsRunProcesss = True
		End If
	Next

	'�g�p�����e��I�u�W�F�N�g����Еt������B
	Set oClassSet = Nothing
	Set oClass = Nothing
	Set oService = Nothing
	Set oLocator = Nothing
End Function

Function KillProcesss(ScriptName)
	Const ProcessName = "wscript.exe"

	'���v���Z�X�͏I���Ώۂ��珜���̂Ŏ��v���Z�XID���擾
	Dim ProcessId
	ProcessId = CurrProcessId

	'WMI�ɂĎg�p����e��I�u�W�F�N�g���`�E��������B
	Dim oClassSet
	Dim oClass
	Dim oLocator
	Dim oService

	'���[�J���R���s���[�^�ɐڑ�����B
	Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set oService = oLocator.ConnectServer
	'�N�G���[������ WQL �ɂĎw�肷��B
	Set oClassSet = oService.ExecQuery("Select * From Win32_Process Where Description=""" & ProcessName & """")
	'�R���N�V��������͂���B
	For Each oClass In oClassSet
		Dim lngPos
		lngPos = InStr(oClass.CommandLine, ScriptName)

		If lngPos <> 0 and oClass.ProcessId <> ProcessId Then
			oClass.Terminate
		End If
	Next

	'�g�p�����e��I�u�W�F�N�g����Еt������B
	Set oClassSet = Nothing
	Set oClass = Nothing
	Set oService = Nothing
	Set oLocator = Nothing
End Function

'���v���Z�XID���擾����(���v���Z�X�͒��ڎ擾�ł��Ȃ��̂Ŏq�v���Z�X������Đe�v���Z�XID���擾����)
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

