Attribute VB_Name = "Main"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

Public MaxCheckBoxConf As Integer  ' USDM�̎d�l��TestLink�Ŏg�p����`�F�b�N�{�b�N�X�̐�
Public CheckBoxSemConf(1 To 5) As String  ' �e�`�F�b�N�{�b�N�X�̈Ӗ�
Public remarksOutConf As Boolean ' ���l���o�͂��邩�ۂ�
Public categoryOutConf As Boolean  ' �J�e�S���[���o�͂��邩�ۂ�
Public sheetNameFirstConf As Boolean  ' �J�e�S���[�񂪂������ꍇ�A�V�[�g���Ƃǂ����D�悵�Ďg����
Public separateReqConf As Boolean ' �v���̗��R�Ɛ������J�X�^�������ɕ����o�͂��邩�ۂ�
Public separateSpcConf As Boolean ' �d�l�̗��R�Ɛ������J�X�^�������ɕ����o�͂��邩�ۂ�
Public separateFVConf As Boolean ' �ړI�@�\(F)�ƌ��ؓ��e(V)�ɂ��āAF���J�X�^�������ɕ����o�͂��邩�ۂ�"

Public Const IDprefix As String = "" ' ����ID�ɑ΂��t�^��prefix
Public Const FVsuffix As String = "" ' ����ID�ɑ΂��AFV�\���̍��ڂɕt�^����ID�Ɏ����I�ɘA������suffix
Public Const ReqSTATUS As String = "D" ' "V"�ɂ��Ă���
Public Const ReqTYPE As String = "3" ' "2"�ɂ��Ă���

Const MaxHeaderRow As Long = 10 ' USDM�̋L�ڂ��n�߂�(�ŏ��̗v�����L�q����)�O�̃w�b�_�̗]�v�ȋL�q���ő�ŉ��s����Ƒz�肷�邩�̒l
Const IniMaxCol As Long = 11 ' USDM�{�̂̃J�����������̏��
Const IniMaxLevel As Long = 100 ' USDM�{�̂̊K�w�������l(���[�v����̏�������ʈ������Ȃ��ōς܂����߂̂���)
Public Const MaxTitle As Long = 75 ' TestLink�̗v�������^�C�g���̒����̏��
Public Const MaxFVtblTitle As Long = 33 ' TestLink�̗v���^�C�g����XML�C���|�[�g�����ɂ����钷���̏��(���ړ��͂Ȃ�����ƒ�������̂����A�A�A)
Public Const MaxTextArea As Long = 235 ' TestLink�̃J�X�^��������TextArea�̒����̏��(255������̂͂�����239��TestLink�̏������ُ�I������)

Const SECTION As String = "1" ' Type: Section
Const USER As String = "2" ' Type: User Requirement Specification
Const System As String = "3" ' Type: System Requirement Specification

Public Const VERIFICATION As String = "Verification" ' ����
Public Const VALIDATION As String = "Validation" ' �Ó����m�F
Public Const IDsuffix As String = "-V" ' �F��d�l�ŗv�����x���Ǝd�l���x����ID���d������̂�h������

''' =================================================================
'''         USDM�V�[�g����͂���XML�ϊ��������w������{�̕���
''' =================================================================
Function createXML(htmlFile, xmlReq, xmlTest, ws, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol As Long) As Boolean ' �V�[�g��XML�ϊ�����
    createXML = False
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim colSpan As Long
    Dim currentCat As String: currentCat = "" ' ���ݍs�̃J�e�S��
    Dim currentReqGrp(1 To 100) As String ' ���ݍs�̗v���O���[�v
    Dim currentSpcGrp As String ' ���ݍs�̎d�l�O���[�v
    Dim reqOrder(1 To 100) As Long ' ���ݍs�̗v���̓���w���ɂ����鏇��
    Dim spcOrder As Long ' ���ݍs�̎d�l�̓���v�����ɂ����鏇��
    Dim currentLevel As Long: currentLevel = 1 ' ���ݍs�̊K�w
    Dim previousLevel As Long: previousLevel = 1000 ' ��O�ɏ��������s�̊K�w
    Dim diffLevel As Long
    Dim specModeFlg As Boolean: specModeFlg = False ' �d�l�o�͒����ۂ�
    Dim content As String
    Dim identifier As String
    Dim baseId As String
    Dim baseIds As String
    Dim reason As String
    Dim description As String
    Dim strSpec(1 To 3) As String
    Dim checkBoxes As String
    Dim remarks As String: remarks = ""
    Dim LogMessage As String ' ���O�o�͗p�ɉ����L�q����Ă���s�Ɣ��肵�������L������
    Dim previousStartCol As Long: previousStartCol = 20 ' ��O�ɏ��������s�����J�����ڂ���n�܂��������L�����Ă����B�������J�e�S���[�J�����͏���
    Dim classification As String ' �e�s�̔F����ʂ��L������
    Dim obj As Object
    
    '�e�m�[�h�p�̕ϐ���錾
    Dim ReqRootElement       As IXMLDOMElement
    Dim targetElement        As IXMLDOMElement ' ���̎��_�ŏ������Ă���v���̃m�[�h
    Dim CurrentParentElement As IXMLDOMElement ' ���̎��_�ŏ������Ă��郌�x���̐e�̃m�[�h
    Dim TestSuiteRootElement As IXMLDOMElement

    If FVtblSCol = 0 Then ' �V�[�g��FV�\�͕K�{�ł���B
        Exit Function
    End If

    ' �������̍��
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
    
    ' <root>�̐���
    Set ReqRootElement = xmlReq.createElement("requirement-specification")
    Call xmlReq.appendChild(ReqRootElement)
    Set CurrentParentElement = ReqRootElement
    ' �e�X�g�̕���<root>�̐���
    Set TestSuiteRootElement = xmlTest.createElement("testsuite")
    Call TestSuiteRootElement.setAttribute("id", "")
    Call TestSuiteRootElement.setAttribute("name", "")
    Call xmlTest.appendChild(TestSuiteRootElement)
    Call InitTestSuite(xmlTest, TestSuiteRootElement)
    
    ' USDM�̕\�S�̂��s�����J��Ԃ��Ȃ����́E�������Ă����{�̕���
    For i = StartRow To MaxRow
        classification = "���m��" ' ���m��ɏ���������
        Print #1, vbTab & "<tr>";
    
        ' USDM�̊e�s����͂��A���ꂪ���̍��ڂ����f���A�l���擾
        For j = Level1Col To MaxCol
            colSpan = 1
            If ws.Cells(i, j).Value <> "" Then ' �����L�q����Ă���
                If classification = "���m��" Then ' �܂��m��ł��Ă��Ȃ��Ƃ�����
                    If IsRequirement(ws.Cells(i, j).Value) Then ' �v���̏ꍇ
                        LogMessage = "�v���̍s�Ƃ��ĔF�����܂���"
                        classification = "�v��"
                    ElseIf IsNintei(ws.Cells(i, j).Value) Then ' �F��d�l�̏ꍇ
                        LogMessage = "�F��d�l�̍s�Ƃ��ĔF�����܂���"
                        classification = "�F��d�l"
                        checkBoxes = Replace(RemoveSpaces(removeCRLF(ws.Cells(i, j).Value)), "�v��", "")
                        If j <> Level1Col Then ' ����������͂P�w�ڂ����ɋ������\���Ƃ���
                            MsgBox ws.Name + "�V�[�g����������" & i & "�s�ڂłP�w�ڈȊO�ɂ͋�����Ȃ��u�v���v�Ɠ����Z���Ƀ`�F�b�N�{�b�N�X��t����u�F��d�l�v�̕\���������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Print #1, "</table>"
                            Print #1, i & "�s�ڂłP�w�ڈȊO�ɂ͋�����Ȃ��u�v���v�Ɠ����Z���Ƀ`�F�b�N�{�b�N�X��t����u�F��d�l�v�̕\���������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                            Close #1
                            Exit Function
                        End If
                    End If
                    If classification = "�v��" Or classification = "�F��d�l" Then ' �v�����F��d�l�̏ꍇ
                        If specModeFlg And j - Level1Col + 1 > currentLevel Then ' �d�l�o�͒��ɊK�w��[������v�����o�Ă���̂�USDM�̋K���ᔽ
                            MsgBox ws.Name + "�V�[�g����������" & i & "�s�ڂŒ��O�́u�d�l�v���u���̎d�l�v�ɂ��Ă��܂��u�v��(�܂��͔F��d�l)�v�������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Print #1, "</table>"
                            Print #1, i & "�s�ڂŒ��O�́u�d�l�v���u���̎d�l�v�ɂ��Ă��܂��u�v��(�܂��͔F��d�l)�v�������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                            Close #1
                            Exit Function
                        End If
                        specModeFlg = False
                        If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ��������J�e�S���[���ς�����Ɣ��f���ĐV���ɃJ�e�S���[���Z�b�g
                            currentCat = ws.Cells(i, 1).Value
                        End If
                        previousLevel = currentLevel
                        currentLevel = j - Level1Col + 1
                        
                        'MsgBox i & "�s�ڂŃJ�����g�̗v���̊K�w���ς��F" & currentLevel - previousLevel
                        diffLevel = currentLevel - previousLevel
                        If diffLevel > 1 Then
                            MsgBox i & "�s�ڂŗv���̊K�w�������Ȃ蕡���w�[���Ȃ�܂����B����͈ᔽ�ł��B"
                        ElseIf diffLevel = 1 Then
                            Set CurrentParentElement = targetElement
                        ElseIf diffLevel < 0 Then
                            For l = currentLevel + 1 To previousLevel
                                Set CurrentParentElement = CurrentParentElement.ParentNode
                                currentReqGrp(l) = "" ' �v���O���[�v������������
                                reqOrder(l) = 0 ' ���̑w�̏��Ԃ�����������
                            Next l
                        End If
                        
                        currentSpcGrp = "" ' �v�����o�Ă����疳�����ɍ��܂ł̎d�l�O���[�v�͖���
                        spcOrder = 0 ' �v�����o�Ă����疳�����Ɏd�l�̏��Ԃ͏�����
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsReason(ws.Cells(i + 1, j + 1).Value) Then ' ���R�̍s�����݂���ꍇ
                            reason = ws.Cells(i + 1, j + 2).Value
                        Else ' ���R�̍s�͑��݂��Ȃ���΂Ȃ�Ȃ�
                            MsgBox ws.Name + "�V�[�g����������USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Print #1, "</table>"
                            Print #1, "USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                            Close #1
                            Exit Function
                        End If
                        If IsDescription(ws.Cells(i + 2, j + 1).Value) Then ' �����̍s�����݂���ꍇ
                            description = ws.Cells(i + 2, j + 2).Value
                        End If
                        If RemarksCol > 0 Then ' ���l�̍s�����݂���ꍇ
                            remarks = ws.Cells(i, RemarksCol).Value
                        End If
                    ElseIf IsReason(ws.Cells(i, j).Value) Then ' ���R����n�܂����ꍇ
                        LogMessage = "���R�̍s�Ƃ��ĔF�����܂���"
                        classification = "���R"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsDescription(ws.Cells(i, j).Value) Then ' ��������n�܂����ꍇ
                        LogMessage = "�����̍s�Ƃ��ĔF�����܂���"
                        classification = "����"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsSpec(ws.Cells(i, j).Value) Then ' �Z���̒������̏�񂩂�d�l�Ɣ��f�����ꍇ
                        LogMessage = "�d�l�̍s�Ƃ��ĔF�����܂���"
                        classification = "�d�l"
                        specModeFlg = True
                        checkBoxes = RemoveSpaces(removeCRLF(ws.Cells(i, j).Value))
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsRequirement(identifier) Then ' �d�l�ԍ�������͂��̉E�ׂ̃Z���Ɂu�v���v������ꍇ
                            LogMessage = "�F��d�l�̍s�Ƃ��ĔF�����܂���"
                            classification = "�F��d�l" ' �F�������߂��@����͐������[��
                            specModeFlg = False
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ��������J�e�S���[���ς�����Ɣ��f���ĐV���ɃJ�e�S���[���Z�b�g
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            ' �F��d�l�͌��X�v���ł��邩��A��������̏����́u�v���v�Ɠ����B���������ׂĂ̗�j�̈�E�ɂ���Ă���
                            previousLevel = currentLevel
                            currentLevel = j + 1 - Level1Col + 1
                            
                            ' MsgBox i & "�s�ڂŃJ�����g�̗v���̊K�w���ς��F" & currentLevel - previousLevel
                            diffLevel = currentLevel - previousLevel
                            If diffLevel > 1 Then
                                MsgBox i & "�s�ڂŗv���̊K�w�������Ȃ蕡���w�[���Ȃ�܂����B����͈ᔽ�ł��B"
                            ElseIf diffLevel = 1 Then
                                Set CurrentParentElement = targetElement
                            ElseIf diffLevel < 0 Then
                                For l = currentLevel + 1 To previousLevel
                                    Set CurrentParentElement = CurrentParentElement.ParentNode
                                    currentReqGrp(l) = "" ' �v���O���[�v������������
                                    reqOrder(l) = 0 ' ���̑w�̏��Ԃ�����������
                                Next l
                            End If
                            currentSpcGrp = "" ' �F��d�l�͗v���ł�����̂ŏo�Ă����疳�����ɍ��܂ł̎d�l�O���[�v�͖���
                            spcOrder = 0 ' �F��d�l�͗v���ł�����̂ŏo�Ă����疳�����Ɏd�l�̏��Ԃ͏�����
                            identifier = ws.Cells(i, j + 1 + 1).Value
                            content = ws.Cells(i, j + 1 + 2).Value
                            If IsReason(ws.Cells(i + 1, j + 1 + 1).Value) Then ' ���R�̍s�����݂���ꍇ
                                reason = ws.Cells(i + 1, j + 1 + 2).Value
                            Else ' ���R�̍s�͑��݂��Ȃ���΂Ȃ�Ȃ�
                                MsgBox ws.Name + "�V�[�g����������USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                                Print #1, "</table>"
                                Print #1, "USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                                Close #1
                                Exit Function
                            End If
                            If IsDescription(ws.Cells(i + 2, j + 1 + 1).Value) Then ' �����̍s�����݂���ꍇ
                                description = ws.Cells(i + 2, j + 1 + 2).Value
                            End If
                        ElseIf IsReason(ws.Cells(i, j + 1).Value) Then ' �E�ׂ̃Z���Ɂu���R�v������ꍇ�@���̃P�[�X�͂����Ǝ��ۂɂ͖���
                            MsgBox ws.Name + "�V�[�g����������USDM�Ƃ��Ă͈Ӗ��s���̃`�F�b�N�{�b�N�X�����Ă��闝�R��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Print #1, "</table>"
                            Print #1, "USDM�Ƃ��Ă͈Ӗ��s���̃`�F�b�N�{�b�N�X�����Ă��闝�R��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                            Close #1
                            Exit Function
                        ElseIf j - Level1Col + 1 <> currentLevel Then ' �d�l�̓J�����g�̗v���̒����ɂȂ���΂Ȃ�Ȃ��B�d�l�̊K�w�\���������Ă͂Ȃ�Ȃ��B
                            MsgBox ws.Name + "�V�[�g����������" & i & "�s�ڂŒ��O�́u�v���v�̒����ɂȂ��u�d�l�v�������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Print #1, "</table>"
                            Print #1, i & "�s�ڂŒ��O�́u�v���v�̒����ɂȂ��u�d�l�v�������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
                            Close #1
                            Exit Function
                        End If
                        If RemarksCol > 0 Then ' ���l�̍s�����݂���ꍇ
                            remarks = ws.Cells(i, RemarksCol).Value
                        End If
                    ElseIf StrComp(Left(ws.Cells(i, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>�ň͂܂�Ă���ꍇ
                        LogMessage = "�O���[�v�̍s�Ƃ��ĔF�����܂���"
                        classification = "�d�l�̃O���[�v" ' �Ƃ܂��͉��肷��
                        content = Mid(ws.Cells(i, j).Value, 2, Len(ws.Cells(i, j).Value) - 2) ' <>�̒������o��
                        If StrComp(Left(content, 1), "<", vbTextCompare) = 0 And StrComp(Right(content, 1), ">", vbTextCompare) = 0 Then ' �Ă�<>�ň͂܂�Ă���ꍇ
                            LogMessage = "�d�l������̍s�Ƃ��ĔF�����܂���"
                            classification = "�d�l�����"
                        ElseIf StrComp(ws.Cells(i + 1, j).Value, "�v��", vbTextCompare) = 0 Then ' �^���̃Z�����v���ł���ꍇ
                            LogMessage = "�v���̃O���[�v�̍s�Ƃ��ĔF�����܂���"
                            classification = "�v���̃O���[�v"
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ��������J�e�S���[���ς�����Ɣ��f���ĐV���ɃJ�e�S���[���Z�b�g
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            currentReqGrp(j - Level1Col + 1) = content
                        ElseIf IsNintei(ws.Cells(i + 1, j).Value) Then ' �^���̃Z�����F��d�l�ł���ꍇ
                            LogMessage = "�v���̃O���[�v�̍s�Ƃ��ĔF�����܂���"
                            classification = "�v���̃O���[�v"
                            If Not sheetNameFirstConf And CategoryCol > 0 And ws.Cells(i, 1).Value <> "" Then ' ��������J�e�S���[���ς�����Ɣ��f���ĐV���ɃJ�e�S���[���Z�b�g
                                currentCat = ws.Cells(i, 1).Value
                            End If
                            currentReqGrp(j - Level1Col + 1) = content
                        End If
                    Else ' �ǂݔ�΂��ׂ��ꍇ
                        LogMessage = "�ǂݔ�΂��ׂ��s�Ƃ��ĔF�����܂���"
                        classification = "�ǂݔ�΂��ׂ�"
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
        
        ' �C���`�LHTML�̃e�[�u���`���ōs�̉�͌��ʂȂǂ��o��
        Print #1, "<td>" & CStr(i) & "�s��" & "</td>";
        Print #1, "<td>" & LogMessage & "</td>";
        Print #1, "<td>" & currentLevel & "�w" & "</td>";
        Print #1, "<td>" & currentSpcGrp & "</td>";
        Print #1, "<td>" & currentReqGrp(1) & ", " & currentReqGrp(2) & ", " & currentReqGrp(3) & ", " & currentReqGrp(4) & ", " & currentReqGrp(5) & ", " & "</td>";
        Print #1, "<td>" & currentCat & "</td>";
        Print #1, "</tr>" & vbCr;
        
        ' �s�̉�͌��ʂɉ�����XML�̃m�[�h���o��
        Select Case classification
            Case "�v��"
                reqOrder(currentLevel) = reqOrder(currentLevel) + 1
                Call appendReqElement(xmlReq, CurrentParentElement, targetElement, identifier, "�v��", checkBoxes, content, reason, description, currentLevel, currentReqGrp(currentLevel), currentCat, remarks, _
                            RemarksCol > 0, _
                            reqOrder(currentLevel), ws, i, FVtblSCol)
                ' �v���̃e�X�g�̓f�t�H���g�őg�����e�X�g���ɏo��
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "�����g����(HAYST�@)�e�X�g")
            ' Case "���R"
            ' Case "����"
            Case "�d�l"
                spcOrder = spcOrder + 1
                Call appendSpecElement(xmlReq, targetElement, _
                        identifier, "�d�l", checkBoxes, content, currentSpcGrp, currentCat, remarks, _
                        RemarksCol > 0, _
                        reqOrder(currentLevel), CStr(spcOrder), ws, i, FVtblSCol)
                ' �d�l�̃e�X�g�̓f�t�H���g�ŒP�@�\�e�X�g���ɏo��
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "�P�@�\�e�X�g")
            Case "�F��d�l"
                reqOrder(currentLevel) = reqOrder(currentLevel) + 1
                Call appendReqElement(xmlReq, CurrentParentElement, targetElement, identifier, "�F��d�l", checkBoxes, content, reason, description, currentLevel, currentReqGrp(currentLevel), currentCat, remarks, _
                            RemarksCol > 0, _
                            reqOrder(currentLevel), ws, i, FVtblSCol)
                ' �F��d�l�̃e�X�g�̓f�t�H���g�ŒP�@�\�e�X�g���Ƒg�����e�X�g���̗����ɏo��
                baseIds = optimizeIds(ws.Cells(i, FVtblSCol + 1).Value)
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, baseIds, "�����g����(HAYST�@)�e�X�g")
                ' �b������B�������ǂ����邩�͂悭�l�������Ȃ���΂Ȃ�Ȃ�
                Call appendTestCaseElement(xmlReq, TestSuiteRootElement, identifier + IDsuffix, "�P�@�\�e�X�g")
            Case "�v���̃O���[�v"
                currentSpcGrp = ""
            Case "�d�l�̃O���[�v"
                currentSpcGrp = content
            ' Case "�O���[�v�����"
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
    MsgBox "�G���[�ԍ�:" & Err.Number
    MsgBox "�G���[���e�F" & Err.description
    MsgBox "�w���v�t�@�C����" & Err.HelpContext
    MsgBox "�v���W�F�N�g���F" & Err.Source
    Resume Next
    Print #1, "</table>"
    Print #1, "�z��O�̃G���[���������Ă��܂������߁A���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B";
    Close #1
End Function

''' ========================================================
'''  USDM���H�����ł���΂ǂ͈̔͂ɋL�ڂ���Ă��邩��͂���
''' ========================================================
Function recognizeUSDMStructure(ws, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) As Boolean
    recognizeUSDMStructure = False
    Dim i As Long
    Dim j As Long
    Dim obj As Object

    Dim lDummy As Long: lDummy = ws.UsedRange.row ' ��x UsedRange ���g���ƍŏI�Z�����␳�����悤��
    MaxRow = ws.Cells.SpecialCells(xlLastCell).row
    MaxCol = ws.Cells.SpecialCells(xlLastCell).Column
    
    If ws.Cells.SpecialCells(xlLastCell).MergeCells Then ' �Z������������ꍇ�ɑΉ����čŏI�Z���̈ʒu���C������
        i = MaxRow
        j = MaxCol
        MaxRow = MaxRow + ws.Cells(i, j).MergeArea.Rows.Count - 1
        MaxCol = MaxCol + ws.Cells(i, j).MergeArea.Columns.Count - 1
    End If
    
    Set obj = ws.Cells.Find("���R", LookAt:=xlWhole) '�܂��́u���R�v�̃Z����T�����Ƃ�USDM�ł��邩�ǂ����̎�|����Ƃ���
    If obj Is Nothing Then
        Exit Function
    End If
    
    For i = 1 To MaxHeaderRow ' ����A�ォ��u�v���v�̃Z����ʂ̕��@�ŒT��(�F��d�l����n�܂�\�����l������)
        For j = 1 To MaxCol
            If IsRequirement(ws.Cells(i, j).Value) Or IsNintei(ws.Cells(i, j).Value) Then ' �v������n�܂����ꍇ�A���邢�́u�v���v�ƃ`�F�b�N�{�b�N�X���܂܂��ꍇ
                If i = obj.row - 1 And j = obj.Column - 1 Then ' ���ꂪ�u���R�v�Z���̍���ɂ������Ȃ�
                    recognizeUSDMStructure = True '���������USDM�ł���Ɣ��f����
                    StartRow = i ' �����Ă܂��������擪�s���Ɖ��肷��
                    If i > 1 Then ' �������������������s��O�ɃO���[�v�̋L�q�����邩������Ȃ�
                        If StrComp(Left(ws.Cells(i - 1, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i - 1, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>�ň͂܂�Ă���ꍇ
                            StartRow = i - 1 ' 1�s�O�ɃO���[�v�̋L�q������Ɣ��f�����̂ŁA������擪�s�Ƃ��ďC������
                        End If
                    End If
                    If i > 2 Then ' ����ɂ������������s��O�ɃO���[�v�̕�����̋L�q�����邩������Ȃ�
                        If StrComp(Left(ws.Cells(i - 2, j).Value, 2), "<<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i - 2, j).Value, 2), ">>", vbTextCompare) = 0 Then ' <<>>�ň͂܂�Ă���ꍇ
                            StartRow = i - 1 ' 2�s�O�ɕ�����̋L�q������Ɣ��f�����̂ŁA������擪�s�Ƃ��ďC������
                        End If
                    End If
                    Level1Col = j ' �J�e�S���[���������ō��J�����i�؍\���̃��[�g�j�ʒu�������ɂ���
                    Exit For
                End If
            End If
        Next j
        If recognizeUSDMStructure Then
            Exit For
        End If
    Next i
    
    Set obj = Nothing
    
    ' �J�e�S���[�񂪂��邩�ǂ������m�F����
    If recognizeUSDMStructure = False Then
        ' MsgBox ws.Name + "�V�[�g��USDM���L�ڂ���Ă�����̂ł͂Ȃ��Ɣ��f���܂����B���̃V�[�g�͏������܂���B"
        Exit Function
    ElseIf Level1Col = 1 Then
        ' MsgBox ws.Name + "�V�[�g�͍ŏ��́u�v���v�܂��͂��̃O���[�v�̕\�L��" & StartRow & "�s" & Level1Col & "��Ɍ����܂����B�J�e�S���[��͑��݂��Ȃ��`���Ɣ��f���A�V�[�g�����J�e�S���[�Ƃ��č̗p���ď������܂��B"
        CategoryCol = 0
    ElseIf Level1Col = 2 Then
        If sheetNameFirstConf Then
            MsgBox ws.Name + "�V�[�g�͍ŏ��́u�v���v�܂��͂��̃O���[�v�̕\�L��" & StartRow & "�s" & Level1Col & "��Ɍ����܂����B1��ڂɃJ�e�S���[���L�q�����`���ł���Ɣ��f����܂��B�������ݒ�ŃV�[�g�����J�e�S���[�Ƃ��Ďg�p����悤�Ɏw�肳��Ă���̂łP��ڂ͎g�p����܂���B)"
        End If
        CategoryCol = 1 ' �ǂ�����g���ɂ���A�񂪂���Ƃ������Ƃ��L������
    Else ' ���̃v���O�����͗]�v�ȗ񂪍��ɂ����Ă������悤�ɏ����Ă���͂������A����e�X�g���ʓ|�Ȃ̂łQ��ȏ゠�����珈������߂�B
        MsgBox ws.Name + "�V�[�g�͍ŏ��́u�v���v�܂��͂��̃O���[�v�̕\�L��" & StartRow & "�s" & Level1Col & "��Ɍ����܂����B�J�e�S���[�ȊO�̗񂪍��ɂ��邱�Ƃ�z�肵�Ă��܂���̂ŁA���̃V�[�g�͏������܂���B"
        Exit Function
    End If
    
    ' ���l�̗񂪂��邩�ǂ������m�F����
    Set obj = ws.Cells.Find("���l��", LookAt:=xlWhole)
    If obj Is Nothing Then
        Set obj = ws.Cells.Find("���l", LookAt:=xlWhole) ' �u���l�v�ł��ǂ����Ƃ�
    End If
    If obj Is Nothing Then ' ����ł��Ȃ��Ȃ���l���͂Ȃ��Ɣ��f
        RemarksCol = 0
        ' MsgBox ws.Name + "�V�[�g�͔��l�����L�ڂ���Ă�����̂ł͂Ȃ��Ɣ��f���܂����B"
    Else
        If obj.row > StartRow Then ' �J�n�s�������Ɂu���l(��)�v������Ƃ������Ƃ͍��ږ��Ƃ��ċL�ڂ��ꂽ�̂ł͂Ȃ���������Ȃ�
            RemarksCol = 0
            MsgBox ws.Name + "�V�[�g�́u���l(��)�v�̋L�ڂ̈ʒu�����̊J�n�s��艺�ɂ���̂ŁA���l�����L�ڂ���Ă�����̂ł͂Ȃ��Ɣ��f���܂����B"
        Else
            RemarksCol = obj.Column
            ' MsgBox "���l���́A" + CStr(RemarksCol) + "��ڂɂ���܂�"
        End If
    End If
    
    ' FV�\�����邩�ǂ����A����Ȃ�΂��̐擪�񂪂ǂ������m�F����
    Set obj = ws.Cells.Find(FItem, LookAt:=xlWhole) '�܂��́u�ړI�@�\�v�̃Z����T�����ƂŊ���FV�\�����邩�ۂ��̎�|����Ƃ���
    If Not obj Is Nothing Then
        FVtblSCol = obj.Cells.Column - 2
    End If
    
    recognizeUSDMStructure = True
End Function

