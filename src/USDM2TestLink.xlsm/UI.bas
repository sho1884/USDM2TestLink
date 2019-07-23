Attribute VB_Name = "UI"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' XML宣言は、マイクロソフトのライブラリが文字コードについて整合的なものを出力しないので、ここで用意した文字列を直接ストリームに出力する
Const XMLDeclaration As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbLf

Const cautionStr As String = "本処理は入力文書フォルダ内のファイルそのものを変更してしまいます。" & vbLf & _
    "（USDMと認識したシートにFV表の列を挿入します。）" & vbLf & _
    "万が一に備えて、入力フォルダ全体のバックアップをとってから実行してください。" & vbLf & _
    "バックアップが未だの場合は、「キャンセル」を押して処理を中止してください。"

Sub startFull()
    start "XML出力処理"
End Sub

Sub startFVtblOnly()
    start "FV表生成のみ"
End Sub

' 最初にこれを呼び出せば良いのだが、上記2つのモードがある
Sub start(mode As String)
    Dim srcPath As String
    Dim destPath As String
    Dim reqPath As String
    Dim testPath As String
    Dim logPath As String
    Dim destBaseName As String
    Dim fileDialog As fileDialog
    Dim fileName As String
    Dim reqFileName As String
    Dim testFileName As String
    Dim logFileName As String
    Dim formType As String
    Dim fvTableStatus As String
    Dim resultMsg As String
    Dim result As Boolean
    Dim ws As Worksheet
    Dim newRow As ListRow
    Dim i As Long
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim StartRow As Long
    Dim Level1Col As Long
    Dim CategoryCol As Long
    Dim RemarksCol As Long
    Dim FVtblSCol As Long
    Dim xmlReq As MSXML2.DOMDocument60
    Dim xmlTest As MSXML2.DOMDocument60
    Dim rc As Integer
    
    rc = MsgBox(cautionStr, vbOKCancel, "警告！")
    If Not rc = vbOK Then
        MsgBox "処理を中止しました"
        Exit Sub
    End If
    
    Worksheets("〔処理結果記録〕").Activate
    Set ws = ActiveSheet
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    Call ClearList
    srcPath = Worksheets("〔XML出力指示＆設定〕").Range("入力パス").Value
    destPath = Worksheets("〔XML出力指示＆設定〕").Range("出力パス").Value
    
    If srcPath = "" Or Dir(srcPath, vbDirectory) = "" Then ' 入力パスが指定されていないか存在しないとき
        srcPath = getPath("処理対象ファイルが格納されているフォルダを選択")
    End If
    If srcPath = "" Then
        MsgBox "存在する入力パスが確定されなかったので処理を終了します。"
        Exit Sub
    End If
    
    If mode = "XML出力処理" Then
        If destPath = "" Or Dir(destPath, vbDirectory) = "" Then ' 出力パスが指定されていないか存在しないとき
            destPath = getPath("処理により生成されるファイルが格納されるフォルダを選択")
        End If
        If destPath = "" Then
            MsgBox "存在する出力パスが確定されなかったので処理を終了します。"
            Exit Sub
        End If
        reqPath = ExportPath(destPath, "要求")
        testPath = ExportPath(destPath, "テスト")
        logPath = ExportPath(destPath, "ログ")
        
        If reqPath = "" Then
            MsgBox "要求情報を出力するパスが確保できなかったので処理を終了します。"
            Exit Sub
        End If
        If testPath = "" Then
            MsgBox "テスト情報を出力するパスが確保できなかったので処理を終了します。"
            Exit Sub
        End If
        If logPath = "" Then
            MsgBox "ログ情報を出力するパスが確保できなかったので処理を終了します。"
            Exit Sub
        End If
    End If
        
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fileName = Dir(srcPath & "\*.xlsx")
    Do While fileName <> ""
        Dim srcBook As Workbook
        Set srcBook = Workbooks.Open(srcPath & "\" & fileName)
        Dim srcSheet As Worksheet
                
        For i = 1 To srcBook.Sheets.Count
            MaxRow = 0
            MaxCol = 0
            StartRow = 0
            Level1Col = 0
            CategoryCol = 0
            RemarksCol = 0
            FVtblSCol = 0
            Set srcSheet = srcBook.Sheets(i)
            destBaseName = srcSheet.Name
            If Not ws.ListObjects("処理記録テーブル").Range.Find(destBaseName, LookAt:=xlWhole) Is Nothing Then ' 既に同じシート名が使われていないか？
                destBaseName = destBaseName & "(" & srcBook.Name & ")"
            End If
            If recognizeUSDMStructure(srcSheet, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) Then
                formType = "USDMと認識"
            Else
                formType = "USDMではない"
            End If
            Set newRow = ws.ListObjects("処理記録テーブル").ListRows.Add
            If formType = "USDMと認識" Then
                fvTableStatus = "生成失敗" ' 初期化
                resultMsg = "処理失敗" ' 初期化
                If FVtblSCol = 0 Then ' FV表の状態を確認する
                    If InsertFVtbl(srcSheet, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) And FVtblSCol > 0 Then
                        fvTableStatus = "今回生成"
                        srcBook.Save
                    End If
                Else
                    fvTableStatus = "既存"
                End If
                If FVtblSCol > 0 And mode = "XML出力処理" Then
                    reqFileName = destBaseName & "-req.xml"
                    testFileName = destBaseName & "-test.xml"
                    logFileName = destBaseName & "-log.html"
                    Set xmlReq = New MSXML2.DOMDocument60
                    Set xmlTest = New MSXML2.DOMDocument60
                    Call createXML(logPath & "\" & logFileName, xmlReq, xmlTest, srcSheet, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol)
                    Call OutputFile(xmlReq, reqPath & "\" & reqFileName)
                    Call OutputFile(xmlTest, testPath & "\" & testFileName)
                    xmlReq.abort
                    xmlTest.abort
                    resultMsg = "処理済み"
                    newRow.Range = Array(fileName, srcSheet.Name, formType, fvTableStatus, resultMsg, reqFileName, testFileName, srcPath, reqPath, testPath, logPath)
                Else
                    newRow.Range = Array(fileName, srcSheet.Name, formType, fvTableStatus, "―", "―", "―", srcPath, "―", "―", "―")
                End If
            Else
                newRow.Range = Array(fileName, srcSheet.Name, formType, "―", "―", "―", "―", srcPath, "―", "―", "―")
            End If
        Next i
        srcBook.Close False
        fileName = Dir()
    Loop
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "処理中にエラーが発生しました。"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub getInputPath()
    Worksheets("〔XML出力指示＆設定〕").Range("入力パス").Value = getPath("処理対象ファイルが格納されているフォルダを選択")
End Sub

Sub getOutputPath()
    Worksheets("〔XML出力指示＆設定〕").Range("出力パス").Value = getPath("処理により生成されるファイルが格納されるフォルダを選択")
End Sub

' ユーザにフォルダを選択させて、そのパスを得る
Function getPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    If fileDialog.Show = -1 Then
        getPath = fileDialog.SelectedItems(1)
    End If
End Function

' 〔処理結果記録〕シートの全データをクリア
Sub ClearList()
    Dim ws As Worksheet
    Worksheets("〔処理結果記録〕").Activate
    Set ws = ActiveSheet
    If ActiveSheet.FilterMode Then
        ws.ShowAllData
    End If
    If Not ws.ListObjects("処理記録テーブル").DataBodyRange Is Nothing Then
        ws.ListObjects("処理記録テーブル").DataBodyRange.ClearContents
    End If
    ws.ListObjects("処理記録テーブル").Resize Range("A1:K2")
End Sub

' 出力する新規のPathを確保する
' basePathの下にcreateFolderNameのフォルダを新規に作成する。
' それが既存の場合は(1)から順に存在しない(N)までを付与した名前を生成して
' 必ず新しい(中が空の)フォルダを生成してそのパスを返す
Function ExportPath(basePath As String, createFolderName As String) As String
    ExportPath = ""
    Dim DirectoryExist, DirectoryPath As String
    Dim i As Long

    ' 指定の既存出力フォルダの中に指定の名前のフォルダーを作る
    If StrComp(Right(basePath, 1), "\", vbTextCompare) <> 0 Then
        basePath = basePath & "\"
    End If

    DirectoryPath = basePath & createFolderName
    DirectoryExist = Dir(DirectoryPath, vbDirectory)

    If DirectoryExist = "" Then
        MkDir DirectoryPath
        ExportPath = DirectoryPath
    Else ' 既存の名前とがかぶったら括弧と番号を付けて新しい名前をつける
        For i = 1 To 1000 ' 実際にはこんなに生成したら管理できないだろう
            DirectoryPath = basePath & createFolderName & "(" & i & ")"
            DirectoryExist = Dir(DirectoryPath, vbDirectory)
            If DirectoryExist = "" Then
                MkDir DirectoryPath
                ExportPath = DirectoryPath
                Exit For
            End If
        Next i
    End If
End Function

' XMLファイルの出力
Function OutputFile(XML As MSXML2.DOMDocument60, fileName As String) As Boolean
    OutputFile = False
    Dim Reader As New SAXXMLReader60
    Dim writer As New MXXMLWriter60

    writer.indent = True
    writer.standalone = True
    
    ' ============ マイクロソフトライブラリの不具合を回避する ============
    'writer.Encoding = "UTF-8"
    writer.Encoding = "shift_jis"
    writer.omitXMLDeclaration = True ' XML宣言は文字コードについて整合的なものが出力されないので、用意した文字列を直接ストリームに出力する
    ' ======== マイクロソフトライブラリの不具合を回避する　おわり ========
    
    Set Reader.contentHandler = writer
    Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8への変換のために使う
    
    Reader.Parse XML.XML
    
    ' ======== TestLink側のバグを回避するための処理 ========
    Dim str As String
    ' TestLink側のバグと思われるが、<status><CDATA>タグ間に改行や空白が入るとエラーになるのでそれを回避する
    str = Replace(writer.output, "<status>＜![CDATA[D]]＞</status>", "<status><![CDATA[D]]></status>")
    ' TestLink側のバグと思われるが、<type><CDATA>タグ間に改行や空白が入るとエラーになるのでそれを回避する
    str = Replace(str, "<type>＜![CDATA[3]]＞</type>", "<type><![CDATA[3]]></type>")
    ' ===== TestLink側のバグを回避するための処理おわり =====
    
    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText XMLDeclaration ' マイクロソフトライブラリの不具合回避
        .LineSeparator = adCRLF
'        .WriteText Replace(writer.output, vbCrLf, vbLf)
        .WriteText Replace(str, vbCrLf, vbLf) ' TestLink側のバグ回避
        .LineSeparator = adLF
        ' BOMを削除する
        Dim byteData() As Byte
        .Position = 0
        .Type = adTypeBinary
        .Position = 3
        byteData = adoSt.Read
        .Close
        .Open
        .Write byteData
    
        .SaveToFile fileName, adSaveCreateOverWrite
        .Close
    End With
    
    OutputFile = True
End Function

' 〔XML出力指示＆設定〕シートの設定値の読み取り
Function GetSetValues() As Boolean
    GetSetValues = False
    Dim i As Long
        
    separateReqConf = False ' 要求の理由や説明の取扱は仕様の中に含める方がデフォルト
    If Worksheets("〔XML出力指示＆設定〕").Range("要求の理由や説明の取扱").Value = "それぞれカスタムフィールドに振り分ける" Then
        separateReqConf = True
    End If
    
    separateSpcConf = False ' 仕様の理由や説明の取扱は仕様の中に含める方がデフォルト
    If Worksheets("〔XML出力指示＆設定〕").Range("仕様の理由や説明の取扱").Value = "それぞれカスタムフィールドに振り分ける" Then
        separateSpcConf = True
    End If
    
    categoryOutConf = True ' カテゴリ情報は出力するのがデフォルト
    sheetNameFirstConf = False ' カテゴリ情報にはセル側にある情報を使うのがデフォルト
    If Worksheets("〔XML出力指示＆設定〕").Range("カテゴリーの取扱").Value = "使わない(出力しない)" Then
        categoryOutConf = False
    ElseIf Worksheets("〔XML出力指示＆設定〕").Range("カテゴリーの取扱").Value = "シート名をカテゴリーとして使用する" Then
        sheetNameFirstConf = True
    End If
    
    If Not Worksheets("〔XML出力指示＆設定〕").Range("グループ名出力の扱い").Value = "グループ名だけの単独ノードは出力しない" Then
        MsgBox "設定項目の「グループ」ですが、現在は「グループ名だけの単独ノードは出力しない」しか実装されていません。"
        Exit Function
    End If
    
    remarksOutConf = True ' 備考欄の情報は出力するのがデフォルト
    If Worksheets("〔XML出力指示＆設定〕").Range("備考欄の取扱").Value = "出力しない" Then
        remarksOutConf = False
        MsgBox "備考欄の情報を出力しないことにした"
    End If
    
    separateFVConf = False ' FV表の目的機能の取扱は目的機能を検証内容に含める方がデフォルト
    If Worksheets("〔XML出力指示＆設定〕").Range("FV表の目的機能の取扱").Value = "目的機能をカスタムフィールドに振り分ける" Then
        separateFVConf = True
    End If

    MaxCheckBoxConf = Worksheets("〔XML出力指示＆設定〕").Range("チェックボックス数").Value
    For i = 1 To MaxCheckBoxConf
        CheckBoxSemConf(i) = RemoveSpaces(removeCRLF(Worksheets("〔XML出力指示＆設定〕").Range("チェックボックスの意味")(i).Value))
        If CheckBoxSemConf(i) = "" Then ' 表の文字列が空だったらデフォルト値にする
            CheckBoxSemConf(i) = "チェックボックス" + CStr(i)
        End If
    Next i
    
    GetSetValues = True
End Function

