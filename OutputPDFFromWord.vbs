'---------------------------------------------------------------------------------
' Microsoft Office Word ファイルをPDFで出力するVBSです。
' Jenkinsに応じた戻り値を返します。
'---------------------------------------------------------------------------------
Option Explicit 
Sub main()
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case ArgCount 
	Case 2
		Dim DocPaths,objshell,PDFPath
		DocPaths = WScript.Arguments(0)
		PDFPath = WScript.Arguments(1)
		StopWordApp
		Set objshell = CreateObject("scripting.filesystemobject")
		If objshell.FolderExists(PDFPath) Then
			If objshell.FolderExists(DocPaths) Then  'Check if the object is a folder
				Dim flag,FileNumber
				flag = 0 
				FileNumber = 0 	
				Dim Folder,DocFiles,DocFile		
				Set Folder = objshell.GetFolder(DocPaths)
				Set DocFiles = Folder.Files
				For Each DocFile In DocFiles  'loop the files in the folder
					FileNumber=FileNumber+1 
					DocPath = DocFile.Path
					If GetWordFile(DocPath) Then  'if the file is Word document, then convert it 
						ConvertWordToPDF DocPath, PDFPath
						flag=flag+1
					End If 	
				Next 
				WScript.Echo "PDF出力が完了しました。" & FileNumber & " 個"
				WScript.Quit "0" ' 正常終了
			Else 
				If GetWordFile(DocPaths) Then  'if the object is a file,then check if the file is a Word document.if that, convert it 
					Dim DocPath
					DocPath = DocPaths
					ConvertWordToPDF DocPath, PDFPath
				Else 
					WScript.Echo "実行時に指定されたWordファイルまたはフォルダが見つかりません。 [" & DocPaths & "]"
					WScript.Quit "1" ' エラー
				End If  
			End If 
		Else 
			WScript.Echo "実行時に指定されたPDF出力先のフォルダが見つかりません。 [ " & PDFPath & "]"
			WScript.Quit "1" ' エラー
		End If
			
	Case Else 
		WScript.Echo "引数は２つ指定してください。 引数１：対象のWordファイルまたはフォルダ。 引数２：出力先のフォルダ。"
		WScript.Quit "1" ' エラー
End Select 
End Sub 

' Wordファイルを指定されたパスにPDFファイルとして出力します。
' DocPath：出力対象のWordファイル
' PDFPath：PDFの出力先パス
Function ConvertWordToPDF(DocPath, PDFPath)
	Dim objshell,ParentFolder,BaseName,wordapp,doc,OutputPath
	Set objshell= CreateObject("scripting.filesystemobject")
	' ParentFolder = objshell.GetParentFolderName(DocPath) 'Get the current folder path
	BaseName = objshell.GetBaseName(DocPath) 'Get the document name
	OutputPath = PDFPath & "\" & BaseName & ".pdf" 
	Set wordapp = CreateObject("Word.application")
	Set doc = wordapp.documents.open(DocPath)
	doc.saveas OutputPath,17
	doc.close
	wordapp.quit
	Set objshell = Nothing 
	WScript.Echo "PDFに出力しました。 doc[" & DocPath & "] PDF[" & OutputPath & "]"
End Function 

'
Function GetWordFile(DocPath) 'Wordファイルを取得します。
	Dim objshell
	Set objshell= CreateObject("scripting.filesystemobject")
	Dim Arrs ,Arr
	Arrs = Array("doc","docx")
	Dim blnIsDocFile,FileExtension
	blnIsDocFile= False 
	FileExtension = objshell.GetExtensionName(DocPath)  'Get the file extension
	For Each Arr In Arrs
		If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then 
			blnIsDocFile= True
			Exit For 
		End If 
	Next 
	GetWordFile = blnIsDocFile
	Set objshell = Nothing 
End Function 

Function StopWordApp '実行中のワードプロセスをkillします。
	Dim strComputer,objWMIService,colProcessList,objProcess 
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	'Get the WinWord.exe
	Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name = 'Winword.exe'")
	For Each objProcess in colProcessList
		'Stop it
		objProcess.Terminate()
	Next
End Function 

Call main 