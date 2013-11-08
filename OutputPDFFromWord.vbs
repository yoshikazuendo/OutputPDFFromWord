'/*
'* Copyright 2013 Yoshikazu Endo
'* 
'* Microsoft Office Word ファイルをPDFで出力するVBSです。
'* PDFへの出力は、WordオブジェクトのSaveAsメソッドを使用しています。
'* また、このメイン関数は、Jenkinsに応じた戻り値を返します。
'* 
'* Tips:
'* http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx
'* 
'*/
Option Explicit 

'***************************************************************************************
' Purpose：メイン関数です。引数を２つ指定する必要があります。 
' Inputs ：引数１：PDFに出力する対象のWordファイルまたはフォルダのフルパス 
'          引数２：PDFの出力先フルパス 
' Returns："0"：正常終了 それ以外：異常終了 
'***************************************************************************************
Sub main()
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case ArgCount 
    Case 2
        Dim docPaths, objshell, pdfPath
        docPaths = WScript.Arguments(0)
        pdfPath = WScript.Arguments(1)
        StopWordApp
        Set objshell = CreateObject("scripting.filesystemobject")
        If objshell.FolderExists(pdfPath) Then
            If objshell.FolderExists(docPaths) Then
                ' フォルダとみなして処理する。 
                Dim outputCount
                outputCount = 0
                Dim Folder, DocFiles, DocFile
                Set Folder = objshell.GetFolder(docPaths)
                Set DocFiles = Folder.Files
                For Each DocFile In DocFiles
                    DocPath = DocFile.Path
                    If IsWordFile(DocPath) Then
                        OutputPDF DocPath, pdfPath
                        outputCount = outputCount + 1
                    End If
                Next 
                WScript.Echo "PDFの出力が完了しました。" & outputCount & " ファイル"
                WScript.Quit "0" ' 正常終了 
            Else 
                ' ファイルとみなして処理する。 
                If IsWordFile(docPaths) Then
                    Dim DocPath
                    DocPath = docPaths
                    OutputPDF DocPath, pdfPath
                    WScript.Echo "PDFの出力が完了しました。"
                    WScript.Quit "0" ' 正常終了 
                Else 
                    WScript.Echo "実行時に指定されたWordファイルまたはフォルダが見つかりません。 [" & docPaths & "]"
                    WScript.Quit "1" ' エラー 
                End If  
            End If 
        Else 
            WScript.Echo "実行時に指定されたPDF出力先のフォルダが見つかりません。 [ " & pdfPath & "]"
            WScript.Quit "1" ' エラー 
        End If
            
    Case Else 
        WScript.Echo "引数は２つ指定してください。 引数１：対象のWordファイルまたはフォルダ。 引数２：出力先のフォルダ。"
        WScript.Quit "1" ' エラー 
End Select 
End Sub 

'***************************************************************************************
' Purpose：実行中のWordプロセスを終了します。 
' Inputs ：なし 
' Returns：なし 
'***************************************************************************************
Function StopWordApp
    Dim strComputer, objWMIService, winwordProccesses, winwordProccess
    strComputer = "."
    ' Winword.exeを取得する。 
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set winwordProccesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'Winword.exe'")
    ' Winword.exeを終了する。
    For Each winwordProccess in winwordProccesses
        winwordProccess.Terminate()
    Next

    Set objWMIService = Nothing
End Function 

'***************************************************************************************
' Purpose：Wordファイルかどうかを返します。 
' Inputs ：Wordファイルのフルパス 
' Returns：True：Wordファイルである。 False：Wordファイルではない。 
'***************************************************************************************
Function IsWordFile(DocPath)
    Dim objshell, extensions, extension, isWord, fileExtension
    Set objshell= CreateObject("scripting.filesystemobject")
    extensions = Array("doc","docx")
    isWord= False
    ' ファイル名から拡張子を取得する。 
    fileExtension = objshell.GetExtensionName(DocPath)

    ' Wordファイルかどうか？ 
    For Each extension In extensions
        If InStr(UCase(fileExtension), UCase(extension)) <> 0 Then 
            isWord= True
            Exit For 
        End If 
    Next 
    IsWordFile = isWord
    Set objshell = Nothing 
End Function 

'***************************************************************************************
' Purpose：Wordファイルの内容をPDFとして出力します。 
' Inputs ：docPath：Wordファイルのフルパス 
'        ：pdfPath：PDFの主力先フルパス 
' Returns：なし 
'***************************************************************************************
Function OutputPDF(docPath, pdfPath)
Dim objshell, fileName, wordapp, doc, outputPath
    Set objshell= CreateObject("scripting.filesystemobject")
    ' 拡張子を除くファイル名を取得する。 
    fileName = objshell.GetBaseName(docPath)
    outputPath = pdfPath & "\" & fileName & ".pdf" 

    ' Wordを開いてPDFで出力する。 
    Set wordapp = CreateObject("Word.application")
    Set doc = wordapp.documents.open(docPath)
    ' 参考情報：Office 2007 はアドイン、2010は標準でsaveAsがPDFをサポートしている。 
    Const wdFormatPDF = 17
    doc.SaveAs outputPath, wdFormatPDF
    doc.close
    wordapp.quit
    Set objshell = Nothing 
    WScript.Echo "PDFに出力しました。 doc[" & docPath & "] PDF[" & outputPath & "]"
End Function 

Call main 