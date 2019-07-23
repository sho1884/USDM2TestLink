Attribute VB_Name = "Sample"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

''' ========================
'''   例示のためのレコード
''' ========================
' 特に推奨できる形式でもないが、、、
Public Const haystsample As String = "<p>FL表(参考：記述形式の例)</p>" & vbLf & _
"" & vbLf & _
"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:500px"">" & vbLf & _
"   <tbody>" & vbLf & _
"       <tr><td><strong>因子</strong></td><td><strong>水準</strong></td></tr>" & vbLf & _
"       <tr><td>因子１</td><td>水準1, 水準2, 水準3</td></tr>" & vbLf & _
"       <tr><td>因子２</td><td>水準1, 水準2, 水準3</td></tr>" & vbLf & _
"   </tbody>" & vbLf & _
"</table>" & vbLf & _
"" & vbLf & _
"<p>&nbsp;</p>" & vbLf & _
""
Public Const SampleSummary As String = _
"<p>前提条件の通りに入力/設定して動作させ、期待値通りになるかを確認します。</p>" & vbLf & _
"<p>（属人的判断を防ぐには、期待値や確認方法も具体的に示します。）</p>"

Public Const SamplePreconditions As String = _
"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""width:500px""><tbody>" & vbLf & _
"    <tr><td><strong>因子</strong></td><td><strong>水準</strong></td></tr>" & vbLf & _
"    <tr><td>因子1</td><td>因子1の値1</td></tr>" & vbLf & _
"    <tr><td>因子2</td><td>因子2の値1</td></tr></tbody></table>" & vbLf & _
"" & vbLf & _
"<p>&nbsp;</p>"

