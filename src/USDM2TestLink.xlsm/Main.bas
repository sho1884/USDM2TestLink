Attribute VB_Name = "Main"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

Public MaxCheckBoxConf As Integer  ' USDMの仕様とTestLinkで使用するチェックボックスの数
Public CheckBoxSemConf(1 To 5) As String  ' 各チェックボックスの意味
Public remarksOutConf As Boolean ' 備考を出力するか否か
Public categoryOutConf As Boolean  ' カテゴリーを出力するか否か
Public sheetNameFirstConf As Boolean  ' カテゴリー列があった場合、シート名とどちらを優先して使うか
Public separateReqConf As Boolean ' 要求の理由と説明をカスタム属性に分離出力するか否か
Public separateSpcConf As Boolean ' 仕様の理由と説明をカスタム属性に分離出力するか否か
Public separateFVConf As Boolean ' 目的機能(F)と検証内容(V)について、Fをカスタム属性に分離出力するか否か"

Public Const IDprefix As String = "" ' 元のIDに対し付与すprefix
Public Const FVsuffix As String = "" ' 元のIDに対し、FV表側の項目に付与するIDに自動的に連結するsuffix
Public Const ReqSTATUS As String = "D" ' "V"にしていた
Public Const ReqTYPE As String = "3" ' "2"にしていた

Const MaxHeaderRow As Long = 10 ' USDMの記載を始める(最初の要求を記述する)前のヘッダの余計な記述が最大で何行あると想定するかの値
Const IniMaxCol As Long = 11 ' USDM本体のカラム数調査の上限
Const IniMaxLevel As Long = 100 ' USDM本体の階層数初期値(ループ初回の処理を特別扱いしないで済ますためのもの)
Public Const MaxTitle As Long = 75 ' TestLinkの要求文書タイトルの長さの上限
Public Const MaxFVtblTitle As Long = 33 ' TestLinkの要求タイトルのXMLインポート処理における長さの上限(直接入力ならもっと長く入るのだが、、、)
Public Const MaxTextArea As Long = 235 ' TestLinkのカスタム属性のTextAreaの長さの上限(255が上限のはずだが239でTestLinkの処理が異常終了する)

Const SECTION As String = "1" ' Type: Section
Const USER As String = "2" ' Type: User Requirement Specification
Const System As String = "3" ' Type: System Requirement Specification

Public Const VERIFICATION As String = "Verification" ' 検証
Public Const VALIDATION As String = "Validation" ' 妥当性確認
Public Const IDsuffix As String = "-V" ' 認定仕様で要求レベルと仕様レベルでIDが重複するのを防ぐため

''' =================================================================
'''         USDMシートを解析してXML変換処理を指示する本体部分
''' =================================================================
Function createXML(htmlFile, xmlReq, xmlTest, ws, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol As Long) As Boolean ' シートをXML変換する
    createXML = False
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim colSpan As Long
    Dim currentCat As String: currentCat = "" ' 現在行のカテゴリ
    Dim currentReqGrp(1 To 100) As String ' 現在行の要求グループ
    Dim currentSpcGrp As String ' 現在行の仕様グループ
    Dim reqOrder(1 To 100) As Long ' 現在行の要求の同一層内における順番
    Dim spcOrder As Long ' 現在行の仕様の同一要求下における順番
    Dim currentLevel As Long: currentLevel = 1 ' 現在行の階層
    Dim previousLevel As Long: previousLevel = 1000 ' 一つ前に処理した行の階層
    Dim diffLevel As Long
    Dim specModeFlg As Boolean: specModeFlg = False ' 仕様出力中か否か
    Dim content As String
    Dim identifier As String
    Dim baseId As String
    Dim baseIds As String
    Dim reason As String
    Dim description As String
    Dim strSpec(1 To 3) As String
    Dim checkBoxes As String
    Dim remarks As String: remarks = ""
    Dim LogMessage As String ' ログ出力用に何が記述されている行と判定したかを記憶する
    Dim previousStartCol As Long: previousStartCol = 20 ' 一つ前に処理した行が何カラム目から始まったかを記憶しておく。ただしカテゴリーカラムは除く
    Dim classification As String ' 各行の認識種別を記憶する
    Dim obj As Object
    
    '各ノード用の変数を宣言
    Dim ReqRootElement       As IXMLDOMElement
    Dim targetElement        As IXMLDOMElement ' その時点で処理している要求のノード
    Dim CurrentParentElement As IXMLDOMElement ' その時点で処理しているレベルの親のノード
    Dim TestSuiteRootElement As IXMLDOMElement

    If FVtblSCol = 0 Then ' シートにFV表は必須である。
        Exit Function
    End If

    ' 初期化の作業
    currentCat = ""
    currentSpcGrp = ""
    For i = 1 To IniMaxLevel
        currentReqGrp(i) = ""
    Next
    spcOrder = 0
    For i = 1 To IniMaxLevel
        reqOrder(i) = 0
    Next
    
    If CategoryCol = 0 Or sheetNameFirstConf Then
        currentCat = ws.Name
    End If
    
    On Error GoTo ErrorTrap
    Open htmlFile For Output As #1
    Print #1, "<table border=""1"">"
    
    ' <root>の生成
    Set ReqRootElement = xmlReq.createElement("requirement-specification")
    Call xmlReq.appendChild(ReqRootElement)
    Set CurrentParentElement = ReqRootElement
    ' テストの方も<root>の生成
    Set TestSuiteRootElement = xmlTest.createElement("testsuite")
    Call TestSuiteRootElement.setAttribute("id", "")
    Call TestSuiteRootElement.setAttribute("name", "")
    Call xmlTest.appendChild(TestSuiteRootElement)
    Call InitTestSuite(xmlTest, TestSuiteRootElement)
    
    ' USDMの表全体を行数分繰り返しながら解析・処理していく本体部分
    For i = StartRow To MaxRow
        classification = "未確定" ' 未確定に初期化する
        Print #1, vbTab & "<tr>";
    
        ' USDMの各行を解析し、それが何の項目か判断し、値を取得
        For j = Level1Col To MaxCol
            colSpan = 1
            If ws.Cells(i, j).Value <> "" Then ' 何か記述されている
                If classification = "未確定" Then ' まだ確定できていないときだけ
                    If IsRequirement(ws.Cells(i, j).Value) Then ' 要求の場合
                        LogMessage = "要求の行として認識しました"
                        classification = "要求"
                    ElseIf IsNintei(ws.Cells(i, j).Value) Then ' 認定仕様の場合
                        LogMessage = "認定仕様の行として認識しました"
                        classification = "認定仕様"
                        checkBoxes = Replace(RemoveSpaces(removeCRLF(ws.Cells(i, j).Value)), "要求", "")
                        If j <> Level1Col Then ' ただしこれは１層目だけに許される表現とする
                            MsgBox ws.Name + "シートを処理中に" & i & "行目で１層目以外には許されない「要求」と同じセルにチェックボックスを付ける「認定仕様」の表現を見つけましたので、このシートの以降の処理を中止します。"
                            Print #1, "</table>"
                            Print #1, i & "行目で１層目以外には許されない「要求」と同じセルにチェックボックスを付ける「認定仕様」の表現を見つけましたので、このシートの以降の処理を中止します。";
                            Close #1
                            Exit Function
                        End If
                    End If
                    If classification = "要求" Or classification = "認定仕様" Then ' 要求か認定仕様の場合
                        If specModeFlg And j - Level1Col + 1 > currentLevel Then ' 仕様出力中に階層を深くする要求が出てくるのはUSDMの規程違反
                            MsgBox ws.Name + "シートを処理中に" & i & "行目で直前の「仕様」を「裸の仕様」にしてしまう「要求(または認定仕様)」を見つけましたので、このシートの以降の処理を中止します。"
                            Print #1, "</table>"
                            Print #1, i & "行目で直前の「仕様」を「裸の仕様」にしてしまう「要求(または認定仕様)」を見つけましたので、このシートの以降の処理を中止します。";
                            Close #1
                            Exit Function
                        End If
                        specModeFlg = False
                        If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ここからカテゴリーが変わったと判断して新たにカテゴリーをセット
                            currentCat = ws.Cells(i, 1).Value
                        End If
                        previousLevel = currentLevel
                        currentLevel = j - Level1Col + 1
                        
                        'MsgBox i & "行目でカレントの要求の階層が変わる：" & currentLevel - previousLevel
                        diffLevel = currentLevel - previousLevel
                        If diffLevel > 1 Then
                            MsgBox i & "行目で要求の階層がいきなり複数層深くなりました。これは違反です。"
                        ElseIf diffLevel = 1 Then
                            Set CurrentParentElement = targetElement
                        ElseIf diffLevel < 0 Then
                            For l = currentLevel + 1 To previousLevel
                                Set CurrentParentElement = CurrentParentElement.ParentNode
                                currentReqGrp(l) = "" ' 要求グループを初期化する
                                reqOrder(l) = 0 ' この層の順番を初期化する
                            Next l
                        End If
                        
                        currentSpcGrp = "" ' 要求が出てきたら無条件に今までの仕様グループは無効
                        spcOrder = 0 ' 要求が出てきたら無条件に仕様の順番は初期化
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsReason(ws.Cells(i + 1, j + 1).Value) Then ' 理由の行が存在する場合
                            reason = ws.Cells(i + 1, j + 2).Value
                        Else ' 理由の行は存在しなければならない
                            MsgBox ws.Name + "シートを処理中にUSDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。"
                            Print #1, "</table>"
                            Print #1, "USDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。";
                            Close #1
                            Exit Function
                        End If
                        If IsDescription(ws.Cells(i + 2, j + 1).Value) Then ' 説明の行が存在する場合
                            description = ws.Cells(i + 2, j + 2).Value
                        End If
                        If RemarksCol > 0 Then ' 備考の行が存在する場合
                            remarks = ws.Cells(i, RemarksCol).Value
                        End If
                    ElseIf IsReason(ws.Cells(i, j).Value) Then ' 理由から始まった場合
                        LogMessage = "理由の行として認識しました"
                        classification = "理由"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsDescription(ws.Cells(i, j).Value) Then ' 説明から始まった場合
                        LogMessage = "説明の行として認識しました"
                        classification = "説明"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsSpec(ws.Cells(i, j).Value) Then ' セルの中だけの情報から仕様と判断される場合
                        LogMessage = "仕様の行として認識しました"
                        classification = "仕様"
                        specModeFlg = True
                        checkBoxes = RemoveSpaces(removeCRLF(ws.Cells(i, j).Value))
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsRequirement(identifier) Then ' 仕様番号があるはずの右隣のセルに「要求」がある場合
                            LogMessage = "認定仕様の行として認識しました"
                            classification = "認定仕様" ' 認識を改めた　これは正式ルール
                            specModeFlg = False
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ここからカテゴリーが変わったと判断して新たにカテゴリーをセット
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            ' 認定仕様は元々要求であるから、ここからの処理は「要求」と同じ。ただしすべての列がjの一つ右にずれている
                            previousLevel = currentLevel
                            currentLevel = j + 1 - Level1Col + 1
                            
                            ' MsgBox i & "行目でカレントの要求の階層が変わる：" & currentLevel - previousLevel
                            diffLevel = currentLevel - previousLevel
                            If diffLevel > 1 Then
                                MsgBox i & "行目で要求の階層がいきなり複数層深くなりました。これは違反です。"
                            ElseIf diffLevel = 1 Then
                                Set CurrentParentElement = targetElement
                            ElseIf diffLevel < 0 Then
                                For l = currentLevel + 1 To previousLevel
                                    Set CurrentParentElement = CurrentParentElement.ParentNode
                                    currentReqGrp(l) = "" ' 要求グループを初期化する
                                    reqOrder(l) = 0 ' この層の順番を初期化する
                                Next l
                            End If
                            currentSpcGrp = "" ' 認定仕様は要求でもあるので出てきたら無条件に今までの仕様グループは無効
                            spcOrder = 0 ' 認定仕様は要求でもあるので出てきたら無条件に仕様の順番は初期化
                            identifier = ws.Cells(i, j + 1 + 1).Value
                            content = ws.Cells(i, j + 1 + 2).Value
                            If IsReason(ws.Cells(i + 1, j + 1 + 1).Value) Then ' 理由の行が存在する場合
                                reason = ws.Cells(i + 1, j + 1 + 2).Value
                            Else ' 理由の行は存在しなければならない
                                MsgBox ws.Name + "シートを処理中にUSDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。"
                                Print #1, "</table>"
                                Print #1, "USDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。";
                                Close #1
                                Exit Function
                            End If
                            If IsDescription(ws.Cells(i + 2, j + 1 + 1).Value) Then ' 説明の行が存在する場合
                                description = ws.Cells(i + 2, j + 1 + 2).Value
                            End If
                        ElseIf IsReason(ws.Cells(i, j + 1).Value) Then ' 右隣のセルに「理由」がある場合　このケースはきっと実際には無い
                            MsgBox ws.Name + "シートを処理中にUSDMとしては意味不明のチェックボックスがついている理由を" & i & "行に見つけましたので、このシートの以降の処理を中止します。"
                            Print #1, "</table>"
                            Print #1, "USDMとしては意味不明のチェックボックスがついている理由を" & i & "行に見つけましたので、このシートの以降の処理を中止します。";
                            Close #1
                            Exit Function
                        ElseIf j - Level1Col + 1 <> currentLevel Then ' 仕様はカレントの要求の直下になければならない。仕様の階層構造もあってはならない。
                            MsgBox ws.Name + "シートを処理中に" & i & "行目で直前の「要求」の直下にない「仕様」を見つけましたので、このシートの以降の処理を中止します。"
                            Print #1, "</table>"
                            Print #1, i & "行目で直前の「要求」の直下にない「仕様」を見つけましたので、このシートの以降の処理を中止します。";
                            Close #1
                            Exit Function
                        End If
                        If RemarksCol > 0 Then ' 備考の行が存在する場合
                            remarks = ws.Cells(i, RemarksCol).Value
                        End If
                    ElseIf StrComp(Left(ws.Cells(i, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>で囲まれている場合
                        LogMessage = "グループの行として認識しました"
                        classification = "仕様のグループ" ' とまずは仮定する
                        content = Mid(ws.Cells(i, j).Value, 2, Len(ws.Cells(i, j).Value) - 2) ' <>の中を取り出す
                        If StrComp(Left(content, 1), "<", vbTextCompare) = 0 And StrComp(Right(content, 1), ">", vbTextCompare) = 0 Then ' 再び<>で囲まれている場合
                            LogMessage = "仕様分割基準の行として認識しました"
                            classification = "仕様分割基準"
                        ElseIf StrComp(ws.Cells(i + 1, j).Value, "要求", vbTextCompare) = 0 Then ' 真下のセルが要求である場合
                            LogMessage = "要求のグループの行として認識しました"
                            classification = "要求のグループ"
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ここからカテゴリーが変わったと判断して新たにカテゴリーをセット
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            currentReqGrp(j - Level1Col + 1) = content
                        ElseIf IsNintei(ws.Cells(i + 1, j).Value) Then ' 真下のセルが認定仕様である場合
                            LogMessage = "要求のグループの行として認識しました"
                            classification = "要求のグループ"
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ここからカテゴリーが変わったと判断して新たにカテゴリーをセット
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            currentReqGrp(j - Level1Col + 1) = content
                        End If
                    Else ' 読み飛ばすべき場合
                        LogMessage = "読み飛ばすべき行として認識しました"
                        classification = "読み飛ばすべき"
                    End If
                End If
            End If
     
            If ws.Cells(i, j).MergeCells Then
                colSpan = ws.Cells(i, j).MergeArea.Columns.Count
                Print #1, "<td colspan=""" & colSpan & """>" & ws.Cells(i, j).Value & "</td>";
            Else
                Print #1, "<td>" & ws.Cells(i, j).Value & "</td>";
            End If
            j = j + colSpan - 1
        Next j
        
        ' インチキHTMLのテーブル形式で行の解析結果などを出力
        Print #1, "<td>" & CStr(i) & "行目" & "</td>";
        Print #1, "<td>" & LogMessage & "</td>";
        Print #1, "<td>" & currentLevel & "層" & "</td>";
        Print #1, "<td>" & currentSpcGrp & "</td>";
        Print #1, "<td>" & currentReqGrp(1) & ", " & currentReqGrp(2) & ", " & currentReqGrp(3) & ", " & currentReqGrp(4) & ", " & currentReqGrp(5) & ", " & "</td>";
        Print #1, "<td>" & currentCat & "</td>";
        Print #1, "</tr>" & vbCr;
        
        ' 行の解析結果に応じてXMLのノードを出力
        Select Case classification
            Case "要求"
                reqOrder(currentLevel) = reqOrder(currentLevel) + 1
                Call appendReqElement(xmlReq, CurrentParentElement, targetElement, identifier, "要求", checkBoxes, content, reason, description, currentLevel, currentReqGrp(currentLevel), currentCat, remarks, _
                            RemarksCol > 0, _
                            reqOrder(currentLevel), ws, i, FVtblSCol)
                ' 要求のテストはデフォルトで組合せテスト側に出力
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "無則組合せ(HAYST法)テスト")
            ' Case "理由"
            ' Case "説明"
            Case "仕様"
                spcOrder = spcOrder + 1
                Call appendSpecElement(xmlReq, targetElement, _
                        identifier, "仕様", checkBoxes, content, currentSpcGrp, currentCat, remarks, _
                        RemarksCol > 0, _
                        reqOrder(currentLevel), CStr(spcOrder), ws, i, FVtblSCol)
                ' 仕様のテストはデフォルトで単機能テスト側に出力
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "単機能テスト")
            Case "認定仕様"
                reqOrder(currentLevel) = reqOrder(currentLevel) + 1
                Call appendReqElement(xmlReq, CurrentParentElement, targetElement, identifier, "認定仕様", checkBoxes, content, reason, description, currentLevel, currentReqGrp(currentLevel), currentCat, remarks, _
                            RemarksCol > 0, _
                            reqOrder(currentLevel), ws, i, FVtblSCol)
                ' 認定仕様のテストはデフォルトで単機能テスト側と組合せテスト側の両方に出力
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "無則組合せ(HAYST法)テスト")
                ' 暫定実装。ここをどうするかはよく考え直さなければならない
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, identifier + IDsuffix, "単機能テスト")
            Case "要求のグループ"
                currentSpcGrp = ""
            Case "仕様のグループ"
                currentSpcGrp = content
            ' Case "グループ分割基準"
            ' Case Else
        End Select
    Next i
    previousLevel = currentLevel
    currentLevel = 1

    Print #1, "</table>"
    Close #1
    createXML = True
    Exit Function

ErrorTrap:
    MsgBox "エラー番号:" & Err.Number
    MsgBox "エラー内容：" & Err.description
    MsgBox "ヘルプファイル名" & Err.HelpContext
    MsgBox "プロジェクト名：" & Err.Source
    Resume Next
    Print #1, "</table>"
    Print #1, "想定外のエラーが発生してしまったため、このシートの以降の処理を中止します。";
    Close #1
End Function

''' ========================================================
'''  USDMか？そうであればどの範囲に記載されているか解析する
''' ========================================================
Function recognizeUSDMStructure(ws, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) As Boolean
    recognizeUSDMStructure = False
    Dim i As Long
    Dim j As Long
    Dim obj As Object

    Dim lDummy As Long: lDummy = ws.UsedRange.row ' 一度 UsedRange を使うと最終セルが補正されるようだ
    MaxRow = ws.Cells.SpecialCells(xlLastCell).row
    MaxCol = ws.Cells.SpecialCells(xlLastCell).Column
    
    If ws.Cells.SpecialCells(xlLastCell).MergeCells Then ' セル結合がある場合に対応して最終セルの位置を修正する
        i = MaxRow
        j = MaxCol
        MaxRow = MaxRow + ws.Cells(i, j).MergeArea.Rows.Count - 1
        MaxCol = MaxCol + ws.Cells(i, j).MergeArea.Columns.Count - 1
    End If
    
    Set obj = ws.Cells.Find("理由", LookAt:=xlWhole) 'まずは「理由」のセルを探すことでUSDMであるかどうかの手掛かりとする
    If obj Is Nothing Then
        Exit Function
    End If
    
    For i = 1 To MaxHeaderRow ' 一方、上から「要求」のセルを別の方法で探す(認定仕様から始まる可能性も考慮する)
        For j = 1 To MaxCol
            If IsRequirement(ws.Cells(i, j).Value) Or IsNintei(ws.Cells(i, j).Value) Then ' 要求から始まった場合、あるいは「要求」とチェックボックスが含まれる場合
                If i = obj.row - 1 And j = obj.Column - 1 Then ' それが「理由」セルの左上にあったなら
                    recognizeUSDMStructure = True 'もうこれはUSDMであると判断する
                    StartRow = i ' そしてまずここが先頭行だと仮定する
                    If i > 1 Then ' しかしもしかしたら一行手前にグループの記述があるかもしれない
                        If StrComp(Left(ws.Cells(i - 1, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i - 1, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>で囲まれている場合
                            StartRow = i - 1 ' 1行前にグループの記述があると判断されるので、ここを先頭行として修正する
                        End If
                    End If
                    If i > 2 Then ' さらにもしかしたら二行手前にグループの分割基準の記述があるかもしれない
                        If StrComp(Left(ws.Cells(i - 2, j).Value, 2), "<<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i - 2, j).Value, 2), ">>", vbTextCompare) = 0 Then ' <<>>で囲まれている場合
                            StartRow = i - 1 ' 2行前に分割基準の記述があると判断されるので、ここを先頭行として修正する
                        End If
                    End If
                    Level1Col = j ' カテゴリーを除いた最左カラム（木構造のルート）位置をここにする
                    Exit For
                End If
            End If
        Next j
        If recognizeUSDMStructure Then
            Exit For
        End If
    Next i
    
    Set obj = Nothing
    
    ' カテゴリー列があるかどうかを確認する
    If recognizeUSDMStructure = False Then
        ' MsgBox ws.Name + "シートはUSDMが記載されているものではないと判断しました。このシートは処理しません。"
        Exit Function
    ElseIf Level1Col = 1 Then
        ' MsgBox ws.Name + "シートは最初の「要求」またはそのグループの表記を" & StartRow & "行" & Level1Col & "列に見つけました。カテゴリー列は存在しない形式と判断し、シート名をカテゴリーとして採用して処理します。"
        CategoryCol = 0
    ElseIf Level1Col = 2 Then
        If sheetNameFirstConf Then
            MsgBox ws.Name + "シートは最初の「要求」またはそのグループの表記を" & StartRow & "行" & Level1Col & "列に見つけました。1列目にカテゴリーが記述される形式であると判断されます。しかし設定でシート名をカテゴリーとして使用するように指定されているので１列目は使用されません。)"
        End If
        CategoryCol = 1 ' どちらを使うにせよ、列があるということを記憶する
    Else ' このプログラムは余計な列が左にあっても動くように書いてあるはずだが、動作テストが面倒なので２列以上あったら処理をやめる。
        MsgBox ws.Name + "シートは最初の「要求」またはそのグループの表記を" & StartRow & "行" & Level1Col & "列に見つけました。カテゴリー以外の列が左にあることを想定していませんので、このシートは処理しません。"
        Exit Function
    End If
    
    ' 備考の列があるかどうかを確認する
    Set obj = ws.Cells.Find("備考欄", LookAt:=xlWhole)
    If obj Is Nothing Then
        Set obj = ws.Cells.Find("備考", LookAt:=xlWhole) ' 「備考」でも良いことに
    End If
    If obj Is Nothing Then ' それでもないなら備考欄はないと判断
        RemarksCol = 0
        ' MsgBox ws.Name + "シートは備考欄が記載されているものではないと判断しました。"
    Else
        If obj.row > StartRow Then ' 開始行よりも下に「備考(欄)」があるということは項目名として記載されたのではないかもしれない
            RemarksCol = 0
            MsgBox ws.Name + "シートは「備考(欄)」の記載の位置が他の開始行より下にあるので、備考欄が記載されているものではないと判断しました。"
        Else
            RemarksCol = obj.Column
            ' MsgBox "備考欄は、" + CStr(RemarksCol) + "列目にあります"
        End If
    End If
    
    ' FV表があるかどうか、あるならばその先頭列がどこかを確認する
    Set obj = ws.Cells.Find(FItem, LookAt:=xlWhole) 'まずは「目的機能」のセルを探すことで既にFV表があるか否かの手掛かりとする
    If Not obj Is Nothing Then
        FVtblSCol = obj.Cells.Column - 2
    End If
    
    recognizeUSDMStructure = True
End Function

