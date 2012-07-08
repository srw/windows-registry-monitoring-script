Dim functions(56)
functions(0) = "advapi32.dll!RegCloseKey"
functions(1) = "advapi32.dll!RegConnectRegistryW"
functions(2) = "advapi32.dll!RegCopyTreeW"
functions(3) = "advapi32.dll!RegCreateKeyW"
functions(4) = "advapi32.dll!RegCreateKeyExW"
functions(5) = "advapi32.dll!RegCreateKeyTransactedW"
functions(6) = "advapi32.dll!RegDeleteKeyW"
functions(7) = "advapi32.dll!RegDeleteKeyExW"
functions(8) = "advapi32.dll!RegDeleteKeyTransactedW"
functions(9) = "advapi32.dll!RegDeleteKeyValueW"
functions(10) = "advapi32.dll!RegDeleteTreeW"
functions(11) = "advapi32.dll!RegDeleteValueW"
functions(12) = "advapi32.dll!RegDisablePredefinedCacheW"
functions(13) = "advapi32.dll!RegDisablePredefinedCacheExW"
functions(14) = "advapi32.dll!RegDisableReflectionKeyW"
functions(15) = "advapi32.dll!RegEnableReflectionKeyW"
functions(16) = "advapi32.dll!RegEnumKeyW"
functions(17) = "advapi32.dll!RegEnumKeyExW"
functions(18) = "advapi32.dll!RegEnumValueW"
functions(19) = "advapi32.dll!RegFlushKeyW"
functions(20) = "advapi32.dll!RegGetValueW"
functions(21) = "advapi32.dll!RegLoadAppKeyW"
functions(22) = "advapi32.dll!RegLoadKeyW"
functions(23) = "advapi32.dll!RegLoadMUIStringW"
functions(24) = "advapi32.dll!RegNotifyChangeKeyValue"
functions(25) = "advapi32.dll!RegOpenCurrentUserW"
functions(26) = "advapi32.dll!RegOpenKeyW"
functions(27) = "advapi32.dll!RegOpenKeyExW"
functions(28) = "advapi32.dll!RegOpenKeyTransactedW"
functions(29) = "advapi32.dll!RegOpenUserClassesRootW"
functions(30) = "advapi32.dll!RegOverridePredefKeyW"
functions(31) = "advapi32.dll!RegQueryInfoKeyW"
functions(32) = "advapi32.dll!RegQueryMultipleValuesW"
functions(33) = "advapi32.dll!RegQueryReflectionKeyW"
functions(34) = "advapi32.dll!RegQueryValueW"
functions(35) = "advapi32.dll!RegQueryValueExW"
functions(36) = "advapi32.dll!RegReplaceKeyW"
functions(37) = "advapi32.dll!RegRestoreKeyW"
functions(38) = "advapi32.dll!RegSaveKeyW"
functions(39) = "advapi32.dll!RegSaveKeyExW"
functions(40) = "advapi32.dll!RegSetKeyValueW"
functions(41) = "advapi32.dll!RegSetValueW"
functions(42) = "advapi32.dll!RegSetValueExW"
functions(43) = "advapi32.dll!RegUnLoadKeyW"
functions(44) = "kernel32.dll!GetPrivateProfileIntW"
functions(45) = "kernel32.dll!GetPrivateProfileSectionW"
functions(46) = "kernel32.dll!GetPrivateProfileSectionNamesW"
functions(47) = "kernel32.dll!GetPrivateProfileStringW"
functions(48) = "kernel32.dll!GetPrivateProfileStructW"
functions(49) = "kernel32.dll!GetProfileIntW"
functions(50) = "kernel32.dll!GetProfileSectionW"
functions(51) = "kernel32.dll!GetProfileStringW"
functions(52) = "kernel32.dll!WritePrivateProfileSectionW"
functions(53) = "kernel32.dll!WritePrivateProfileStringW"
functions(54) = "kernel32.dll!WritePrivateProfileStructW"
functions(55) = "kernel32.dll!WriteProfileSectionW"
functions(56) = "kernel32.dll!WriteProfileStringW"

Dim spyMgr


Set spyMgr = WScript.CreateObject("DeviareCOM.NktSpyMgr", "spyMgr_")

spyMgr.Initialize

WScript.Echo "NkySpyMgr Initialized"

Dim hooks
Set hooks = CreateObject("Scripting.Dictionary")
Dim func

Dim hook
For Each func in functions
	Set hook = spyMgr.CreateHook(func, 33)
	'hook.Hook True
	hooks.add func, hook
Next

WScript.Echo "Hooks Created"

Dim processes, process

Set processes = spyMgr.Processes()
Set process = processes.First()

Do While Not process Is Nothing
If LCase(process.Name) = "outlook.exe" And process.PlatformBits = 32 Then
'If LCase(process.Name) = "iexplore.exe" And process.PlatformBits = 32 Then
		For Each key in hooks
			'WScript.Echo key
			WScript.Echo "Before hooking " & key & " on iexplore.exe"
			'MsgBox TypeName(hooks.Item(key))
			hooks.Item(key).Attach process, False
			WScript.Echo "After hooking " & key & " on iexplore.exe"
		Next
	End If
	Set process = processes.Next()
Loop

WScript.Echo "Hooks attached to processes"

For Each key in hooks
	hooks.Item(key).Hook True
Next

WScript.Echo "Hooks in True"

Sub spyMgr_OnFunctionCalled(ByVal hook, ByVal proc, ByVal callInfo)
	Dim params, param, log

	log = hook.FunctionName
	Set params = callInfo.Params()
	If Not params Is Nothing Then
		log = log & " #parameters = " & params.Count & " param names: "
		Set param = params.First()
		Do While Not param is Nothing
			log = log & param.Name & " "
			Set param = params.Next()
		Loop

	End If

	WScript.Echo log
End Sub

Dim wshShell
Set wshShell = WScript.CreateObject("WScript.Shell")
wshShell.Popup "Press the OK button to exit this script"
