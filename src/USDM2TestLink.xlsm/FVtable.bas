Attribute VB_Name = "FVtable"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

Const FVtblNCol As Integer = 8 ' FV表部分の全カラム数
Public Const TBidItem As String = "テストベースID(No.)"
Public Const FItem As String = "目的機能(F)"
Const VItem As String = "検証内容(V)"
Public Const TItem As String = "テスト技法(T)"
Public Const PdRItem As String = "市場リスク"
Const PdRItemL As String = "市場リスク" & vbLf & "(プロダクトリスク)"
Public Const PjRItem As String = "技術リスク"
Const PjRItemL As String = "技術リスク" & vbLf & "(プロジェクトリスク)"
Public Const FLFPItem As String = "FLFP"
Const FLFPItemL As String = "FLFP(Factor Level Function Point)"
Public Const VVItem As String = "V&V区分"
Const FVInteriorColorIndex As Integer = 0 ' FV表部分の背景色
Const FVFontColorIndex As Integer = 1 ' FV表部分の文字色

''' ================================
'''          FV表生成処理
''' ================================
' アクティブなシートにFV表を挿入する
' 副作用としてFV表の開始カラムFVtblSColを返す。RemarksColがあれば、その位置が修正される
' また副作用として一行目に項目名行が挿入されるため、一行下にずれることもある
Function InsertFVtbl(ws, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) As Boolean
    InsertFVtbl = False
    Dim FVtblWidth As Variant: FVtblWidth = Array(20, 20, 80, 60, 60, 20, 20, 20)
    Dim FVtblTitle As Variant: FVtblTitle = Array(VVItem, TBidItem, FItem, VItem, TItem, PdRItemL, PjRItemL, FLFPItemL)
    Dim i As Long
    Dim j As Long
    Dim obj As Object
    Dim classification As String
    Dim rowsItem As Long
    Dim colSpan As Long
    Dim identifier As String
    Dim content As String
    Dim reason As String
    Dim description As String
    Dim rc As Integer
    
    If FVtblSCol <> 0 Then
        MsgBox ws.Name & "シートには既に" & FVtblSCol & "カラム目からFV表が存在すると認識されています。二重追加はできません。"
        Exit Function
    End If
    ws.Activate
    

    If RemarksCol = 0 Then
        FVtblSCol = MaxCol + 1 ' 備考欄がない場合はその時点の最終カラムの右隣から
    Else
        FVtblSCol = RemarksCol ' 備考欄がある場合はその位置に挿入する
        For j = 1 To FVtblNCol
            Columns(FVtblSCol).Insert
        Next j
        RemarksCol = FVtblSCol + FVtblNCol ' 備考欄の位置を修正する
    End If

    rc = MsgBox(ws.Parent.Name & "の" & ws.Name & "シートにFV表を、" + CStr(FVtblSCol) + "列目から" + CStr(FVtblSCol + FVtblNCol - 1) + "列目に挿入します", vbOKCancel, "警告！")
    If Not rc = vbOK Then
        MsgBox "処理を中止しました"
        Exit Function
    End If
'    MsgBox ws.Parent.Name & "の" & ws.Name & "シートにFV表を、" + CStr(FVtblSCol) + "列目から" + CStr(FVtblSCol + FVtblNCol - 1) + "列目に挿入します"

    If StartRow < 2 Then ' USDMには項目名行が必ずしもないが、FV表には必ず項目名行が必要なので、その行を確保する
        Rows(1).Insert
        StartRow = StartRow + 1
        MaxRow = MaxRow + 1
    End If

    For j = 0 To FVtblNCol - 1 ' 空のFV表を追加(挿入)
        Columns(FVtblSCol + j).ColumnWidth = FVtblWidth(j)
        Call MergeForce(StartRow - 1, FVtblSCol + j, StartRow - 1, FVtblSCol + j, FVtblTitle(j))
    Next j

    ' USDMの各行を解析しながら処理していく本体部分
    For i = StartRow To MaxRow
        classification = "未確定"
        rowsItem = 1 ' 項目毎の行数はデフォルトで1とする
    
        ' USDMの行を解析して値を取得
        For j = Level1Col To MaxCol
            colSpan = 1
            If ws.Cells(i, j).Value <> "" Then ' 何か記述されている
                If classification = "未確定" Then ' 何の項目が記述されているのか判明しているならもう処理は不要
                    If IsRequirement(ws.Cells(i, j).Value) Then ' 要求の場合
                        classification = "要求"
                    ElseIf IsNintei(ws.Cells(i, j).Value) Then ' 認定仕様の場合
                        classification = "認定仕様"
                        If j <> Level1Col Then ' ただしこれは１層目だけに許される表現とする
                            MsgBox ws.Name + "シートを処理中に" & i & "行目で１層目以外には許されない「要求」と同じセルにチェックボックスを付ける「認定仕様」の表現を見つけましたので、このシートの以降の処理を中止します。"
                            Exit Function
                        End If
                    End If
                    If classification = "要求" Or classification = "認定仕様" Then ' 要求と認定仕様の場合
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsReason(ws.Cells(i + 1, j + 1).Value) Then ' 理由の行が存在する場合
                            reason = ws.Cells(i + 1, j + 2).Value
                        Else ' 理由の行は存在しなければならない
                            MsgBox ws.Name + "シートを処理中にUSDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。"
                            Exit Function
                        End If
                        rowsItem = 2
                        If IsDescription(ws.Cells(i + 2, j + 1).Value) Then ' 説明の行が存在する場合
                            rowsItem = 3
                            description = ws.Cells(i + 2, j + 2).Value
                        End If
                    ElseIf IsReason(ws.Cells(i, j).Value) Then ' 理由から始まった場合
                        classification = "理由"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsDescription(ws.Cells(i, j).Value) Then ' 説明から始まった場合
                        classification = "説明"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsSpec(ws.Cells(i, j).Value) Then ' セルの中だけの情報から仕様と判断される場合
                        classification = "仕様"
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsRequirement(identifier) Then ' 仕様番号があるはずの右隣のセルに「要求」がある場合
                            classification = "認定仕様" ' 認識を改めた　これは正式ルール
                            ' 認定仕様は元々要求であるから、ここからの処理は「要求」と同じ。ただしすべての列がjの一つ右にずれている
                            identifier = ws.Cells(i, j + 1 + 1).Value
                            content = ws.Cells(i, j + 1 + 2).Value
                            If IsReason(ws.Cells(i + 1, j + 1 + 1).Value) Then ' 理由の行が存在する場合
                                reason = ws.Cells(i + 1, j + 1 + 2).Value
                            Else ' 理由の行は存在しなければならない
                                MsgBox ws.Name + "シートを処理中にUSDMで推奨されない理由の行を持たない認定仕様(要求)を" & i & "行に見つけましたので、このシートの以降の処理を中止します。"
                                Exit Function
                            End If
                            rowsItem = 2
                            If IsDescription(ws.Cells(i + 2, j + 1 + 1).Value) Then ' 説明の行が存在する場合
                                rowsItem = 3
                                description = ws.Cells(i + 2, j + 1 + 2).Value
                            End If
                        End If
                    ElseIf StrComp(Left(ws.Cells(i, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>で囲まれている場合
                        classification = "仕様のグループ" ' とまずは仮定する
                        content = Mid(ws.Cells(i, j).Value, 2, Len(ws.Cells(i, j).Value) - 2) ' <>の中を取り出す
                        If StrComp(Left(content, 1), "<", vbTextCompare) = 0 And StrComp(Right(content, 1), ">", vbTextCompare) = 0 Then ' 再び<>で囲まれている場合
                            classification = "仕様分割基準"
                        ElseIf StrComp(ws.Cells(i + 1, j).Value, "要求", vbTextCompare) = 0 Then ' 真下のセルが要求である場合
                            classification = "要求のグループ"
                        ElseIf IsNintei(ws.Cells(i + 1, j).Value) Then ' 真下のセルが認定仕様である場合
                            classification = "要求のグループ"
                        End If
                    Else ' 読み飛ばすべき場合
                        classification = "読み飛ばすべき"
                    End If
                Else
                    Exit For ' 横方向の解析をやめる
                End If
            End If
     
            If ws.Cells(i, j).MergeCells Then
                colSpan = ws.Cells(i, j).MergeArea.Columns.Count
            End If
            j = j + colSpan - 1
        Next j
        
        ' 行の解析結果に応じてFV表の下書き内容を出力
        Select Case classification
            Case "要求"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Validation", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[理由転記：" & reason & "]" & vbLf & "[要求転記：" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "例えば因子を列挙します"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "例えば組合せテスト, シナリオテスト"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "未評価", "未評価,大,中,小")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "未評価", "未評価,高,中,低")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            ' Case "理由"
            ' Case "説明"
            Case "仕様"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Verification", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[仕様転記：" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "例えば因子を列挙します"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "例えばデシジョンテーブル"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "未評価", "未評価,大,中,小")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "未評価", "未評価,高,中,低")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            Case "認定仕様"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Validation", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[理由転記：" & reason & "]" & vbLf & "[要求転記：" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "例えば因子を列挙します"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "例えば組合せテスト"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "未評価", "未評価,大,中,小")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "未評価", "未評価,高,中,低")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            ' Case "要求のグループ"
            ' Case "仕様のグループ"
            ' Case "グループ分割基準"
            ' Case Else
        End Select
    Next i
    
    MsgBox "FV表の挿入は正常に最後まで処理されました。"
    InsertFVtbl = True
End Function

Function MergeForce(ByVal row1 As Long, ByVal column1 As Long, ByVal row2 As Long, ByVal column2 As Long, ByVal str As String, Optional listStr As String = "")
    Range(Cells(row1, column1), Cells(row2, column2)).Select
    Selection.Merge
    Selection.Value = str
    Selection.Interior.ColorIndex = FVInteriorColorIndex
    Selection.Font.ColorIndex = FVFontColorIndex
    Selection.Borders.LineStyle = xlContinuous
    Selection.WrapText = True
    If (Len(listStr) > 0) Then
        With Selection.VALIDATION
            .Delete
            .Add Type:=xlValidateList, _
                 Operator:=xlEqual, _
                 Formula1:=listStr
        End With
    End If
End Function

