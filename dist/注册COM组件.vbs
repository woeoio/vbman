With CreateObject("Scripting.FileSystemObject")
	Dim BaseDir : BaseDir = .GetParentFolderName(WScript.ScriptFullName)
	Dim BasePart : BasePart = Split(BaseDir, ":")
	Dim CmdStr : CmdStr = " /k cd """ & BaseDir & """&" & BasePart(0) & ":"
	
	Dim ComVBRC
	Dim RegCmd
	
	ComVBRC = "VBMAN.dll"
	RegCmd = "Regsvr32 " & ComVBRC
	CmdStr = CmdStr & "&" & RegCmd
	
	'ComVBRC = "NTSVC.ocx"
	'RegCmd = "Regsvr32 " & ComVBRC
	'CmdStr = CmdStr  & "&" & RegCmd
	
	CreateObject("Shell.Application").ShellExecute "Cmd.exe", """" & CmdStr & """", "", "runas", 1
	'MsgBox CmdStr, , "×¢²á³É¹¦"
End With