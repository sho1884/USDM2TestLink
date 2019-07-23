Attribute VB_Name = "Utility"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 要求かどうかを判定します。
Function IsRequirement(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "要求", vbTextCompare) = 0 Then
        IsRequirement = True
    Else
        IsRequirement = False
    End If
End Function

' 認定仕様かどうかを判定します。(セル内だけからの判断なので、「要求」の隣のセルにチェックボックスがついている場合のことは判断できない。)
Function IsNintei(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If str <> "" Then
        If Len(str) > Len("要求") Then
            str = RemoveCheckBox(str)
            If str = "" Then ' チェックボックスだけが記載されていたということになる場合
                IsNintei = False
            ElseIf StrComp(str, "要求", vbTextCompare) = 0 Then ' チェックボックスも入っていたことになる
                IsNintei = True
            Else ' 「要求」とチェックボックス以外の何かが記載されていた場合
                IsNintei = False
            End If
        Else
            IsNintei = False ' 長さ２以下で認定仕様は表せない。「要求」のみはダメ
        End If
    Else
        IsNintei = False ' 最初から空なら認定仕様項目ではない
    End If
End Function

' 仕様の項目か(つまりはチェックボックスだけしかないセルか)どうかを判定します。
' (セル内だけからの判断なので、隣のセルが「要求」なら認定仕様の可能性もある。)
Function IsSpec(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If str <> "" Then
        str = RemoveCheckBox(str)
        If str = "" Then ' チェックボックスだけが記載されていたということになる場合
            IsSpec = True
        Else ' チェックボックス以外の何かが記載されていた場合
            IsSpec = False
        End If
    Else
        IsSpec = False ' 最初から空なら仕様項目ではない
    End If
End Function

' 理由かどうかを判定します。
Function IsReason(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "理由", vbTextCompare) = 0 Then
        IsReason = True
    Else
        IsReason = False
    End If
End Function

' 説明かどうかを判定します。
Function IsDescription(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "説明", vbTextCompare) = 0 Then
        IsDescription = True
    Else
        IsDescription = False
    End If
End Function

' チェックボックスの文字を取り除いた文字列を返す
Function RemoveCheckBox(str As String) As String
    str = Replace(str, "□", "")
    str = Replace(str, "■", "")
    str = Replace(str, ChrW(9745), "") 'レでチェックされているボックス
    RemoveCheckBox = Replace(str, ChrW(9746), "") '×でチェックされているボックス
End Function

' 改行の文字を取り除いた文字列を返す
Function removeCRLF(str As String) As String
    str = Replace(str, vbCr, "")
    removeCRLF = Replace(str, vbLf, "")
End Function

' 空白の文字を取り除いた文字列を返す
Function RemoveSpaces(str As String) As String
    str = Replace(str, " ", "")
    RemoveSpaces = Replace(str, "　", "")
End Function

' HTML表現における特殊文字の回避
Function EscapeHTML(ByVal str As String) As String
  str = Replace(str, "&", "&amp;", Compare:=vbBinaryCompare) ' これを最初に置換しないと下記置換で現れるものに影響する
  str = Replace(str, "<", "&lt;", Compare:=vbBinaryCompare)
  str = Replace(str, ">", "&gt;", Compare:=vbBinaryCompare)
  str = Replace(str, """", "&quot;", Compare:=vbBinaryCompare)
  str = Replace(str, "'", "&#039;", Compare:=vbBinaryCompare)
  EscapeHTML = str
End Function

' 要求・理由・説明を合わせてHTML表現に変換（特殊文字の回避と改行文字をHTMLタグに変換するだけ）
Function concatToHTML(ByVal content As String, ByVal reason As String, ByVal description As String) As String
    Dim concatStr As String
    content = EscapeHTML(content)
    content = Replace(content, vbCr, "")
    content = Replace(content, vbLf, "<br />" & vbLf)
    reason = EscapeHTML(reason)
    reason = Replace(reason, vbCr, "")
    reason = Replace(reason, vbLf, "<br />" & vbLf)
    description = EscapeHTML(description)
    description = Replace(description, vbCr, "")
    description = Replace(description, vbLf, "<br />" & vbLf)
    concatStr = "<h1>要求</h1>" & vbLf
    concatStr = concatStr & "<p>" & content & "</p>" & vbLf
    concatStr = concatStr & "<h1>理由</h1>" & vbLf
    concatStr = concatStr & "<p>" & reason & "</p>" & vbLf
    If Not description = "" Then
        concatStr = concatStr & "<h1>説明</h1>" & vbLf
        concatStr = concatStr & "<p>" & description & "</p>" & vbLf
    End If
    concatToHTML = concatStr
End Function

' 目的機能(F)・検証内容(V)を合わせてHTML表現に変換（特殊文字の回避と改行文字をHTMLタグに変換するだけ）
Function concatFVToHTML(ByVal fContent As String, ByVal vContent As String) As String
    Dim concatStr As String
    fContent = EscapeHTML(fContent)
    fContent = Replace(fContent, vbCr, "")
    fContent = Replace(fContent, vbLf, "<br />" & vbLf)
    vContent = EscapeHTML(vContent)
    vContent = Replace(vContent, vbCr, "")
    vContent = Replace(vContent, vbLf, "<br />" & vbLf)
    concatStr = "<h1>目的機能(F):</h1>" & vbLf
    concatStr = concatStr & "<p>" & fContent & "</p>" & vbLf
    concatStr = concatStr & "<h1>検証内容(V):</h1>" & vbLf
    concatStr = concatStr & "<p>" & vContent & "</p>" & vbLf
    concatFVToHTML = concatStr
End Function

' HTML表現に変換（特殊文字の回避と改行文字をHTMLタグに変換するだけ）
Function toHTML(ByVal str As String) As String
    str = EscapeHTML(str)
    str = Replace(str, vbCr, "")
    toHTML = Replace(str, vbLf, "<br />" + vbLf)
End Function

' CDATA表現に変換（TestLinkのカスタマイズTextArea用）
Function toCDATA(ByVal str As String) As String
    str = Replace(str, vbCr, "")
    If Len(str) > MaxTextArea Then
        Debug.Print "元の長さ: " & Len(str)
        str = Left(str, MaxTextArea - Len("★長さが上限を超えたので超過分を削除しました★") - 1)
        toCDATA = "★長さが上限を超えたので超過分を削除しました★" & vbLf & str
        Debug.Print "修正済みの長さ: " & Len(toCDATA)
    Else
        toCDATA = str
    End If
End Function

' チェックボックス列を意味する値の列に変換
Function CheckBox2ValStr(str As String) As String
    str = RemoveSpaces(removeCRLF(str))
    Dim retStr As String: retStr = ""
    Dim i As Long
    For i = 1 To Len(str)
        If i > 5 Then Exit For ' チェックボックスの数は５までと決めているので残りは無視
        If StrComp(Mid(str, i, 1), "□", vbTextCompare) <> 0 Then ' チェック入りであるとき
            retStr = retStr + CheckBoxSemConf(i) + "|"
        End If
    Next i
    Dim l As Long: l = Len(retStr)
    If l > 0 Then retStr = Left(retStr, l - 1) ' 最後の"|"を取り除く
    CheckBox2ValStr = retStr
End Function

' USDMの要求または仕様の本文を切り取って要件仕様(requirement spec)のタイトル項目の文字列を作る
Function makeTitle(ByVal str As String) As String
    makeTitle = curtail(str, MaxTitle)
End Function

' FV表のVの本文を切り取って要件文書(requirement)のタイトル項目の文字列を作る
Function makeFVtblTitle(ByVal str As String) As String
'    str = EscapeHTML(str)
    makeFVtblTitle = curtail(str, MaxFVtblTitle)
End Function

' 1行に詰め込む。入りきらないなら空行でない一番上の1行を採用。それもだめなら強制的に切る
Function curtail(str As String, maxLen As Long) As String
    Dim retStr As String: retStr = removeCRLF(str)
    If Len(retStr) <= maxLen Then ' 詰め込んで入りきるならそのまま採用
        curtail = retStr
        Exit Function
    End If
    Do While StrComp(Left(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Left(str, 1), vbCr, vbTextCompare) = 0
        str = Mid(str, 2) ' 最初の１文字を捨てる
    Loop
    Dim pCr As Long: pCr = InStr(str, vbCr)
    Dim pLf As Long: pLf = InStr(str, vbLf)
    Dim p As Long: p = 0
    If pCr = 0 Then ' このコードは明らかに冗長だがわかり易さのために敢えてそうした
        If pLf = 0 Then ' 両方見つからなかったときは
            p = 0
        Else ' LFだけが入っていたということになるので、その位置
            p = pLf
        End If
    Else
        If pLf = 0 Then ' CRだけが入っていたということになるので
            p = pCr
        Else ' 両方入っていたときには先に出現した方を取る
            If pCr > pLf Then
                p = pLf
            Else
                p = pCr
            End If
        End If
    End If
    If p = 0 Or p - 1 > maxLen Then ' 改行が入っていないか、あるいは入っていても制限を超えてしまうなら
        retStr = Left(str, maxLen - 1) + "…" ' 強制的に切るしか無い
    Else ' 改行まで(つまり１行目だけ)なら入りきるのでこれをタイトルとして採用
        retStr = Left(str, p - 1)
    End If
    curtail = retStr
End Function

 ' 仕様内の理由や説明を分解する
 ' （理由や説明を一つの項目内で複数回記載することは無いと仮定している）
Function separateSpec(str As String, retStr() As String) As Boolean
    Dim pReason As Long: pReason = InStr(str, "【理由】")
    Dim pDescription As Long: pDescription = InStr(str, "【説明】")
    If pReason > 0 And pDescription > 0 Then ' 【理由】と【説明】の両方がある
        If pReason < pDescription Then ' 【理由】の方が先に出てくるパターン
            retStr(1) = Left(str, pReason - 1)
            retStr(2) = Mid(str, pReason + 4, pDescription - pReason - 4)
            retStr(3) = Mid(str, pDescription + 4)
        Else ' 【説明】の方が先に出てくるパターン
            retStr(1) = Left(str, pDescription - 1)
            retStr(2) = Mid(str, pReason + 4)
            retStr(3) = Mid(str, pDescription + 4, pReason - pDescription - 4)
        End If
    ElseIf pReason > 0 Then ' 既に片方しか無いことがわかっているので【理由】だけ
        retStr(1) = Left(str, pReason - 1)
        retStr(2) = Mid(str, pReason + 4)
        retStr(3) = ""
    ElseIf pReason > 0 Then ' 既に片方しか無いことがわかっているので【説明】だけ
        retStr(1) = Left(str, pDescription - 1)
        retStr(2) = ""
        retStr(3) = Mid(str, pDescription + 4)
    Else ' どちらもないことが確定した
        retStr(1) = str
        retStr(2) = ""
        retStr(3) = ""
    End If
    ' 先頭と末尾の改行を取り去る
    retStr(1) = removeCRLFbothEnds(retStr(1))
    retStr(2) = removeCRLFbothEnds(retStr(2))
    retStr(3) = removeCRLFbothEnds(retStr(3))
    separateSpec = True
End Function

' 文字列の両端に限って改行を取り除く
Function removeCRLFbothEnds(str As String) As String
    Do While StrComp(Left(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Left(str, 1), vbCr, vbTextCompare) = 0
        str = Mid(str, 2) ' 最初の１文字を捨てる
    Loop
    Do While StrComp(Right(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Right(str, 1), vbCr, vbTextCompare) = 0
        str = Left(str, Len(str) - 1) ' 最初の１文字を捨てる
    Loop
    removeCRLFbothEnds = str
End Function

' ベースIDの列からIDを一つずつ取り出す。
Function extractId(ByRef str As String) As String
    Dim strLen As Long
    Dim e As Long
    strLen = Len(str)
    If strLen = 0 Then
        extractId = vbNullString
    Else
        e = InStr(1, str, ",", vbTextCompare) ' 両端がカンマの場合は既に取り除いている
        If e = 0 Then
            extractId = str
            str = ""
        Else
            extractId = Left(str, e - 1) ' だからeが1になることはあり得ない
            If strLen > e Then
                str = Mid(str, e + 1)
            Else
                str = ""
            End If
        End If
    End If
End Function

' ベースIDの列の記述を最適化する。
Function optimizeIds(str As String) As String
    Dim strLen As Long
    Dim priorStrLen As Long
    Dim e As Long
    Dim tmpStr As String
    tmpStr = Replace(str, " ", "") ' 半角スペースと取り除く
    tmpStr = Replace(tmpStr, "　", "") ' 全角スペースと取り除く
    tmpStr = Replace(tmpStr, vbCr, ",") ' 改行をカンマに置き換える
    tmpStr = Replace(tmpStr, vbLf, ",") ' 改行をカンマに置き換える
    tmpStr = removeComment(tmpStr) ' コメント部を取り除いてカンマに置き換える
    priorStrLen = Len(tmpStr)
    tmpStr = Replace(tmpStr, ",,", ",") ' カンマの連続を一つに置き換える
    strLen = Len(tmpStr)
    Do While priorStrLen <> strLen
        priorStrLen = strLen
        tmpStr = Replace(tmpStr, ",,", ",") ' カンマの連続を一つに置き換える
        strLen = Len(tmpStr)
    Loop
    If StrComp(Left(tmpStr, 1), ",", vbTextCompare) = 0 Then ' 左端のカンマを取り除く
        tmpStr = Mid(tmpStr, 2)
    End If
    If StrComp(Right(tmpStr, 1), ",", vbTextCompare) = 0 Then ' 右端のカンマを取り除く
        tmpStr = Left(tmpStr, Len(tmpStr) - 1)
    End If
    optimizeIds = tmpStr
End Function

' ベースIDの列からコメント部を取り除いてカンマに置き換える
Function removeComment(str As String) As String
    Dim strLen As Long
    Dim priorStrLen As Long
    Dim s As Long
    Dim e As Long
    Dim tmpStr As String
    removeComment = str
    priorStrLen = 0
    strLen = Len(removeComment)
    Do While priorStrLen <> strLen
        priorStrLen = strLen
        s = InStr(1, removeComment, "[", vbTextCompare)
        If s > 0 Then ' コメント部があるので取り除く
            e = InStr(s, removeComment, "]", vbTextCompare)
            If e = 0 Then ' コメントの終端が無いので最後までコメントと見做す
                e = strLen
            End If
            removeComment = Replace(removeComment, Mid(removeComment, s, e - s + 1), ",") ' コメントはカンマに置き換える
        End If
        strLen = Len(removeComment)
    Loop
End Function

