Attribute VB_Name = "Utility"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' �v�����ǂ����𔻒肵�܂��B
Function IsRequirement(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "�v��", vbTextCompare) = 0 Then
        IsRequirement = True
    Else
        IsRequirement = False
    End If
End Function

' �F��d�l���ǂ����𔻒肵�܂��B(�Z������������̔��f�Ȃ̂ŁA�u�v���v�ׂ̗̃Z���Ƀ`�F�b�N�{�b�N�X�����Ă���ꍇ�̂��Ƃ͔��f�ł��Ȃ��B)
Function IsNintei(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If str <> "" Then
        If Len(str) > Len("�v��") Then
            str = RemoveCheckBox(str)
            If str = "" Then ' �`�F�b�N�{�b�N�X�������L�ڂ���Ă����Ƃ������ƂɂȂ�ꍇ
                IsNintei = False
            ElseIf StrComp(str, "�v��", vbTextCompare) = 0 Then ' �`�F�b�N�{�b�N�X�������Ă������ƂɂȂ�
                IsNintei = True
            Else ' �u�v���v�ƃ`�F�b�N�{�b�N�X�ȊO�̉������L�ڂ���Ă����ꍇ
                IsNintei = False
            End If
        Else
            IsNintei = False ' �����Q�ȉ��ŔF��d�l�͕\���Ȃ��B�u�v���v�݂̂̓_��
        End If
    Else
        IsNintei = False ' �ŏ������Ȃ�F��d�l���ڂł͂Ȃ�
    End If
End Function

' �d�l�̍��ڂ�(�܂�̓`�F�b�N�{�b�N�X���������Ȃ��Z����)�ǂ����𔻒肵�܂��B
' (�Z������������̔��f�Ȃ̂ŁA�ׂ̃Z�����u�v���v�Ȃ�F��d�l�̉\��������B)
Function IsSpec(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If str <> "" Then
        str = RemoveCheckBox(str)
        If str = "" Then ' �`�F�b�N�{�b�N�X�������L�ڂ���Ă����Ƃ������ƂɂȂ�ꍇ
            IsSpec = True
        Else ' �`�F�b�N�{�b�N�X�ȊO�̉������L�ڂ���Ă����ꍇ
            IsSpec = False
        End If
    Else
        IsSpec = False ' �ŏ������Ȃ�d�l���ڂł͂Ȃ�
    End If
End Function

' ���R���ǂ����𔻒肵�܂��B
Function IsReason(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "���R", vbTextCompare) = 0 Then
        IsReason = True
    Else
        IsReason = False
    End If
End Function

' �������ǂ����𔻒肵�܂��B
Function IsDescription(str As String) As Boolean
    str = RemoveSpaces(removeCRLF(str))
    If StrComp(str, "����", vbTextCompare) = 0 Then
        IsDescription = True
    Else
        IsDescription = False
    End If
End Function

' �`�F�b�N�{�b�N�X�̕�������菜�����������Ԃ�
Function RemoveCheckBox(str As String) As String
    str = Replace(str, "��", "")
    str = Replace(str, "��", "")
    str = Replace(str, ChrW(9745), "") '���Ń`�F�b�N����Ă���{�b�N�X
    RemoveCheckBox = Replace(str, ChrW(9746), "") '�~�Ń`�F�b�N����Ă���{�b�N�X
End Function

' ���s�̕�������菜�����������Ԃ�
Function removeCRLF(str As String) As String
    str = Replace(str, vbCr, "")
    removeCRLF = Replace(str, vbLf, "")
End Function

' �󔒂̕�������菜�����������Ԃ�
Function RemoveSpaces(str As String) As String
    str = Replace(str, " ", "")
    RemoveSpaces = Replace(str, "�@", "")
End Function

' HTML�\���ɂ�������ꕶ���̉��
Function EscapeHTML(ByVal str As String) As String
  str = Replace(str, "&", "&amp;", Compare:=vbBinaryCompare) ' ������ŏ��ɒu�����Ȃ��Ɖ��L�u���Ō������̂ɉe������
  str = Replace(str, "<", "&lt;", Compare:=vbBinaryCompare)
  str = Replace(str, ">", "&gt;", Compare:=vbBinaryCompare)
  str = Replace(str, """", "&quot;", Compare:=vbBinaryCompare)
  str = Replace(str, "'", "&#039;", Compare:=vbBinaryCompare)
  EscapeHTML = str
End Function

' �v���E���R�E���������킹��HTML�\���ɕϊ��i���ꕶ���̉���Ɖ��s������HTML�^�O�ɕϊ����邾���j
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
    concatStr = "<h1>�v��</h1>" & vbLf
    concatStr = concatStr & "<p>" & content & "</p>" & vbLf
    concatStr = concatStr & "<h1>���R</h1>" & vbLf
    concatStr = concatStr & "<p>" & reason & "</p>" & vbLf
    If Not description = "" Then
        concatStr = concatStr & "<h1>����</h1>" & vbLf
        concatStr = concatStr & "<p>" & description & "</p>" & vbLf
    End If
    concatToHTML = concatStr
End Function

' �ړI�@�\(F)�E���ؓ��e(V)�����킹��HTML�\���ɕϊ��i���ꕶ���̉���Ɖ��s������HTML�^�O�ɕϊ����邾���j
Function concatFVToHTML(ByVal fContent As String, ByVal vContent As String) As String
    Dim concatStr As String
    fContent = EscapeHTML(fContent)
    fContent = Replace(fContent, vbCr, "")
    fContent = Replace(fContent, vbLf, "<br />" & vbLf)
    vContent = EscapeHTML(vContent)
    vContent = Replace(vContent, vbCr, "")
    vContent = Replace(vContent, vbLf, "<br />" & vbLf)
    concatStr = "<h1>�ړI�@�\(F):</h1>" & vbLf
    concatStr = concatStr & "<p>" & fContent & "</p>" & vbLf
    concatStr = concatStr & "<h1>���ؓ��e(V):</h1>" & vbLf
    concatStr = concatStr & "<p>" & vContent & "</p>" & vbLf
    concatFVToHTML = concatStr
End Function

' HTML�\���ɕϊ��i���ꕶ���̉���Ɖ��s������HTML�^�O�ɕϊ����邾���j
Function toHTML(ByVal str As String) As String
    str = EscapeHTML(str)
    str = Replace(str, vbCr, "")
    toHTML = Replace(str, vbLf, "<br />" + vbLf)
End Function

' CDATA�\���ɕϊ��iTestLink�̃J�X�^�}�C�YTextArea�p�j
Function toCDATA(ByVal str As String) As String
    str = Replace(str, vbCr, "")
    If Len(str) > MaxTextArea Then
        Debug.Print "���̒���: " & Len(str)
        str = Left(str, MaxTextArea - Len("������������𒴂����̂Œ��ߕ����폜���܂�����") - 1)
        toCDATA = "������������𒴂����̂Œ��ߕ����폜���܂�����" & vbLf & str
        Debug.Print "�C���ς݂̒���: " & Len(toCDATA)
    Else
        toCDATA = str
    End If
End Function

' �`�F�b�N�{�b�N�X����Ӗ�����l�̗�ɕϊ�
Function CheckBox2ValStr(str As String) As String
    str = RemoveSpaces(removeCRLF(str))
    Dim retStr As String: retStr = ""
    Dim i As Long
    For i = 1 To Len(str)
        If i > 5 Then Exit For ' �`�F�b�N�{�b�N�X�̐��͂T�܂łƌ��߂Ă���̂Ŏc��͖���
        If StrComp(Mid(str, i, 1), "��", vbTextCompare) <> 0 Then ' �`�F�b�N����ł���Ƃ�
            retStr = retStr + CheckBoxSemConf(i) + "|"
        End If
    Next i
    Dim l As Long: l = Len(retStr)
    If l > 0 Then retStr = Left(retStr, l - 1) ' �Ō��"|"����菜��
    CheckBox2ValStr = retStr
End Function

' USDM�̗v���܂��͎d�l�̖{����؂����ėv���d�l(requirement spec)�̃^�C�g�����ڂ̕���������
Function makeTitle(ByVal str As String) As String
    makeTitle = curtail(str, MaxTitle)
End Function

' FV�\��V�̖{����؂����ėv������(requirement)�̃^�C�g�����ڂ̕���������
Function makeFVtblTitle(ByVal str As String) As String
'    str = EscapeHTML(str)
    makeFVtblTitle = curtail(str, MaxFVtblTitle)
End Function

' 1�s�ɋl�ߍ��ށB���肫��Ȃ��Ȃ��s�łȂ���ԏ��1�s���̗p�B��������߂Ȃ狭���I�ɐ؂�
Function curtail(str As String, maxLen As Long) As String
    Dim retStr As String: retStr = removeCRLF(str)
    If Len(retStr) <= maxLen Then ' �l�ߍ���œ��肫��Ȃ炻�̂܂܍̗p
        curtail = retStr
        Exit Function
    End If
    Do While StrComp(Left(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Left(str, 1), vbCr, vbTextCompare) = 0
        str = Mid(str, 2) ' �ŏ��̂P�������̂Ă�
    Loop
    Dim pCr As Long: pCr = InStr(str, vbCr)
    Dim pLf As Long: pLf = InStr(str, vbLf)
    Dim p As Long: p = 0
    If pCr = 0 Then ' ���̃R�[�h�͖��炩�ɏ璷�����킩��Ղ��̂��߂Ɋ����Ă�������
        If pLf = 0 Then ' ����������Ȃ������Ƃ���
            p = 0
        Else ' LF�����������Ă����Ƃ������ƂɂȂ�̂ŁA���̈ʒu
            p = pLf
        End If
    Else
        If pLf = 0 Then ' CR�����������Ă����Ƃ������ƂɂȂ�̂�
            p = pCr
        Else ' ���������Ă����Ƃ��ɂ͐�ɏo�������������
            If pCr > pLf Then
                p = pLf
            Else
                p = pCr
            End If
        End If
    End If
    If p = 0 Or p - 1 > maxLen Then ' ���s�������Ă��Ȃ����A���邢�͓����Ă��Ă������𒴂��Ă��܂��Ȃ�
        retStr = Left(str, maxLen - 1) + "�c" ' �����I�ɐ؂邵������
    Else ' ���s�܂�(�܂�P�s�ڂ���)�Ȃ���肫��̂ł�����^�C�g���Ƃ��č̗p
        retStr = Left(str, p - 1)
    End If
    curtail = retStr
End Function

 ' �d�l���̗��R������𕪉�����
 ' �i���R���������̍��ړ��ŕ�����L�ڂ��邱�Ƃ͖����Ɖ��肵�Ă���j
Function separateSpec(str As String, retStr() As String) As Boolean
    Dim pReason As Long: pReason = InStr(str, "�y���R�z")
    Dim pDescription As Long: pDescription = InStr(str, "�y�����z")
    If pReason > 0 And pDescription > 0 Then ' �y���R�z�Ɓy�����z�̗���������
        If pReason < pDescription Then ' �y���R�z�̕�����ɏo�Ă���p�^�[��
            retStr(1) = Left(str, pReason - 1)
            retStr(2) = Mid(str, pReason + 4, pDescription - pReason - 4)
            retStr(3) = Mid(str, pDescription + 4)
        Else ' �y�����z�̕�����ɏo�Ă���p�^�[��
            retStr(1) = Left(str, pDescription - 1)
            retStr(2) = Mid(str, pReason + 4)
            retStr(3) = Mid(str, pDescription + 4, pReason - pDescription - 4)
        End If
    ElseIf pReason > 0 Then ' ���ɕЕ������������Ƃ��킩���Ă���̂Ły���R�z����
        retStr(1) = Left(str, pReason - 1)
        retStr(2) = Mid(str, pReason + 4)
        retStr(3) = ""
    ElseIf pReason > 0 Then ' ���ɕЕ������������Ƃ��킩���Ă���̂Ły�����z����
        retStr(1) = Left(str, pDescription - 1)
        retStr(2) = ""
        retStr(3) = Mid(str, pDescription + 4)
    Else ' �ǂ�����Ȃ����Ƃ��m�肵��
        retStr(1) = str
        retStr(2) = ""
        retStr(3) = ""
    End If
    ' �擪�Ɩ����̉��s����苎��
    retStr(1) = removeCRLFbothEnds(retStr(1))
    retStr(2) = removeCRLFbothEnds(retStr(2))
    retStr(3) = removeCRLFbothEnds(retStr(3))
    separateSpec = True
End Function

' ������̗��[�Ɍ����ĉ��s����菜��
Function removeCRLFbothEnds(str As String) As String
    Do While StrComp(Left(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Left(str, 1), vbCr, vbTextCompare) = 0
        str = Mid(str, 2) ' �ŏ��̂P�������̂Ă�
    Loop
    Do While StrComp(Right(str, 1), vbLf, vbTextCompare) = 0 Or StrComp(Right(str, 1), vbCr, vbTextCompare) = 0
        str = Left(str, Len(str) - 1) ' �ŏ��̂P�������̂Ă�
    Loop
    removeCRLFbothEnds = str
End Function

' �x�[�XID�̗񂩂�ID��������o���B
Function extractId(ByRef str As String) As String
    Dim strLen As Long
    Dim e As Long
    strLen = Len(str)
    If strLen = 0 Then
        extractId = vbNullString
    Else
        e = InStr(1, str, ",", vbTextCompare) ' ���[���J���}�̏ꍇ�͊��Ɏ�菜���Ă���
        If e = 0 Then
            extractId = str
            str = ""
        Else
            extractId = Left(str, e - 1) ' ������e��1�ɂȂ邱�Ƃ͂��蓾�Ȃ�
            If strLen > e Then
                str = Mid(str, e + 1)
            Else
                str = ""
            End If
        End If
    End If
End Function

' �x�[�XID�̗�̋L�q���œK������B
Function optimizeIds(str As String) As String
    Dim strLen As Long
    Dim priorStrLen As Long
    Dim e As Long
    Dim tmpStr As String
    tmpStr = Replace(str, " ", "") ' ���p�X�y�[�X�Ǝ�菜��
    tmpStr = Replace(tmpStr, "�@", "") ' �S�p�X�y�[�X�Ǝ�菜��
    tmpStr = Replace(tmpStr, vbCr, ",") ' ���s���J���}�ɒu��������
    tmpStr = Replace(tmpStr, vbLf, ",") ' ���s���J���}�ɒu��������
    tmpStr = removeComment(tmpStr) ' �R�����g������菜���ăJ���}�ɒu��������
    priorStrLen = Len(tmpStr)
    tmpStr = Replace(tmpStr, ",,", ",") ' �J���}�̘A������ɒu��������
    strLen = Len(tmpStr)
    Do While priorStrLen <> strLen
        priorStrLen = strLen
        tmpStr = Replace(tmpStr, ",,", ",") ' �J���}�̘A������ɒu��������
        strLen = Len(tmpStr)
    Loop
    If StrComp(Left(tmpStr, 1), ",", vbTextCompare) = 0 Then ' ���[�̃J���}����菜��
        tmpStr = Mid(tmpStr, 2)
    End If
    If StrComp(Right(tmpStr, 1), ",", vbTextCompare) = 0 Then ' �E�[�̃J���}����菜��
        tmpStr = Left(tmpStr, Len(tmpStr) - 1)
    End If
    optimizeIds = tmpStr
End Function

' �x�[�XID�̗񂩂�R�����g������菜���ăJ���}�ɒu��������
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
        If s > 0 Then ' �R�����g��������̂Ŏ�菜��
            e = InStr(s, removeComment, "]", vbTextCompare)
            If e = 0 Then ' �R�����g�̏I�[�������̂ōŌ�܂ŃR�����g�ƌ���
                e = strLen
            End If
            removeComment = Replace(removeComment, Mid(removeComment, s, e - s + 1), ",") ' �R�����g�̓J���}�ɒu��������
        End If
        strLen = Len(removeComment)
    Loop
End Function

