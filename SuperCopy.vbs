''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                            '
'             SuperCopy                                                                      '
'   ��;�����������տ���Զ�����ӹ����У���ֹ�����ļ��ĳ����£������ļ�����                   '
'   ���ߣ�������                                                                             '
'   ���䣺5592440@qq.com                                                                     '
'   2020��5��11��                                                                            '
'                                                                                            '
'   Enjoy yourself!                                                                          '
'                                                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Const TypeBinary = 1
Const ForReading = 1, ForWriting = 2, ForAppending = 8
 
' getting file from args (no checks!)
Dim arguments, runType, inFile, outFile, inByteArray, base64Encoded, base64Decoded, outByteArray,text, fileName
 
Set arguments = WScript.Arguments
'MsgBox arguments.length
If arguments.length = 0 Then
	'��װ
	Install
Else
	runType = arguments(0)
	'MsgBox runType
	If runType = 1 Then
		'Copy
		inFile = arguments(1)
		fileName = Right(inFile,Len(inFile) - InStrRev(inFile,"\"))
		'MsgBox fileName
		inByteArray = readBytes(inFile)
		base64Encoded = encodeBase64(inByteArray)
		text = base64Encoded & "^" & fileName
		'MsgBox text
		SetClipboardText text
		'MsgBox "�ѿ�����������"
	ElseIf runType = 2 Then
		'Paste
		outFile = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
		text = GetClipboardText
		'MsgBox text
		Wscript.echo "Base64 encoded: " + text
		base64Encoded = Left(text,InStr(text,"^") - 1)
		Wscript.echo "Base64 encoded: " + base64Encoded
		fileName = Trim(Right(text,Len(text) - InStr(text,"^")))
		fileName = Replace(fileName,vbCrLf,"")
		'MsgBox outFile+fileName
		outFile = outFile & "\" & fileName
		base64Decoded = decodeBase64(base64Encoded)
		writeBytes outFile, base64Decoded
	Else
		MsgBox '����Ĳ���'
	End If
End If



Private Function readBytes(file)
  Dim inStream
  ' ADODB stream object used
  Set inStream = WScript.CreateObject("ADODB.Stream")
  ' open with no arguments makes the stream an empty container 
  inStream.Open
  inStream.type= TypeBinary
  inStream.LoadFromFile(file)
  readBytes = inStream.Read()
End Function
 
Private Function encodeBase64(bytes)
  Dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
End Function
 
Private Function decodeBase64(base64)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set encoded String, get bytes
  EL.Text = base64
  decodeBase64 = EL.NodeTypedValue
end function
 
Private Sub writeBytes(file, bytes)
  Dim binaryStream
  Set binaryStream = CreateObject("ADODB.Stream")
  binaryStream.Type = TypeBinary
  'Open the stream and write binary data
  binaryStream.Open
  binaryStream.Write bytes
  'Save binary data to disk��
  binaryStream.SaveToFile file, ForWriting
End Sub

Private Sub SetClipboardText(Text)   'д����Ϣ�����а�
	Dim WshShell, oExec, oIn
	Set WshShell = CreateObject("WScript.Shell")
	Set oExec = WshShell.Exec("clip")
	Set oIn = oExec.stdIn	
	oIn.WriteLine text
	oIn.Close
End Sub

Private Function GetClipboardText()   '���ж�ȡ����Ϣ
	Dim text
	text = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
	GetClipboardText = text
End Function

Private Function Install()
	MsgBox "�����װ��ɺ��Ҽ��˵���δ���֡�������������������ճ�����˵����롾ʹ�ù���ԱȨ�����С�cmd.exe,����vbs��ȫ·�����С�",vbOKOnly,"��ʼ��װ"
	Dim curDir,path, ws,fso,f,str
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	curDir = fso.GetFolder(".").Path 
	path = curDir + "\reg.reg"
	Set f = fso.OpenTextFile( path, ForAppending, True)
	f.WriteLine("Windows Registry Editor Version 5.00")
	f.WriteBlankLines(1)
	f.WriteLine("[HKEY_CLASSES_ROOT\*\shell\SuperCopy]")
	f.WriteLine("@=""��������""")
	f.WriteBlankLines(1)
	f.WriteLine("[HKEY_CLASSES_ROOT\*\shell\SuperCopy\command]")
	str = "@=""CScript.exe " & WScript.ScriptFullName & " 1 %1"""
	str = Replace(str,"\","\\")
	f.WriteLine(str)
	f.WriteBlankLines(1)
	f.WriteLine("[HKEY_CLASSES_ROOT\Directory\Background\shell\SuperPaste]")
	f.WriteLine("@=""����ճ��""")
	f.WriteBlankLines(1)
	f.WriteLine("[HKEY_CLASSES_ROOT\Directory\Background\shell\SuperPaste\command]")
	str = "@=""CScript.exe " & WScript.ScriptFullName & " 2"""
	str = Replace(str,"\","\\")
	f.WriteLine(str)
	f.Close
	
	Set ws = CreateObject("WScript.Shell")
	ws.Run "regedit /s " & path
	
	fso.DeleteFile(path)
	
	MsgBox "�Ҽ������������������ļ����ƣ�" & vbCrLf & "�Ҽ�������ճ���������ļ�ճ��" & vbCrLf & "��ӭʹ�ã�" & vbCrLf & "�������뷴����5592440@qq.com",vbOKOnly, "��װ�ɹ�"
End Function
