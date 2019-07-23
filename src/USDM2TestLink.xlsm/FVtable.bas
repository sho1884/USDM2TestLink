Attribute VB_Name = "FVtable"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

Const FVtblNCol As Integer = 8 ' FV�\�����̑S�J������
Public Const TBidItem As String = "�e�X�g�x�[�XID(No.)"
Public Const FItem As String = "�ړI�@�\(F)"
Const VItem As String = "���ؓ��e(V)"
Public Const TItem As String = "�e�X�g�Z�@(T)"
Public Const PdRItem As String = "�s�ꃊ�X�N"
Const PdRItemL As String = "�s�ꃊ�X�N" & vbLf & "(�v���_�N�g���X�N)"
Public Const PjRItem As String = "�Z�p���X�N"
Const PjRItemL As String = "�Z�p���X�N" & vbLf & "(�v���W�F�N�g���X�N)"
Public Const FLFPItem As String = "FLFP"
Const FLFPItemL As String = "FLFP(Factor Level Function Point)"
Public Const VVItem As String = "V&V�敪"
Const FVInteriorColorIndex As Integer = 0 ' FV�\�����̔w�i�F
Const FVFontColorIndex As Integer = 1 ' FV�\�����̕����F

''' ================================
'''          FV�\��������
''' ================================
' �A�N�e�B�u�ȃV�[�g��FV�\��}������
' ����p�Ƃ���FV�\�̊J�n�J����FVtblSCol��Ԃ��BRemarksCol������΁A���̈ʒu���C�������
' �܂�����p�Ƃ��Ĉ�s�ڂɍ��ږ��s���}������邽�߁A��s���ɂ���邱�Ƃ�����
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
        MsgBox ws.Name & "�V�[�g�ɂ͊���" & FVtblSCol & "�J�����ڂ���FV�\�����݂���ƔF������Ă��܂��B��d�ǉ��͂ł��܂���B"
        Exit Function
    End If
    ws.Activate
    

    If RemarksCol = 0 Then
        FVtblSCol = MaxCol + 1 ' ���l�����Ȃ��ꍇ�͂��̎��_�̍ŏI�J�����̉E�ׂ���
    Else
        FVtblSCol = RemarksCol ' ���l��������ꍇ�͂��̈ʒu�ɑ}������
        For j = 1 To FVtblNCol
            Columns(FVtblSCol).Insert
        Next j
        RemarksCol = FVtblSCol + FVtblNCol ' ���l���̈ʒu���C������
    End If

    rc = MsgBox(ws.Parent.Name & "��" & ws.Name & "�V�[�g��FV�\���A" + CStr(FVtblSCol) + "��ڂ���" + CStr(FVtblSCol + FVtblNCol - 1) + "��ڂɑ}�����܂�", vbOKCancel, "�x���I")
    If Not rc = vbOK Then
        MsgBox "�����𒆎~���܂���"
        Exit Function
    End If
'    MsgBox ws.Parent.Name & "��" & ws.Name & "�V�[�g��FV�\���A" + CStr(FVtblSCol) + "��ڂ���" + CStr(FVtblSCol + FVtblNCol - 1) + "��ڂɑ}�����܂�"

    If StartRow < 2 Then ' USDM�ɂ͍��ږ��s���K�������Ȃ����AFV�\�ɂ͕K�����ږ��s���K�v�Ȃ̂ŁA���̍s���m�ۂ���
        Rows(1).Insert
        StartRow = StartRow + 1
        MaxRow = MaxRow + 1
    End If

    For j = 0 To FVtblNCol - 1 ' ���FV�\��ǉ�(�}��)
        Columns(FVtblSCol + j).ColumnWidth = FVtblWidth(j)
        Call MergeForce(StartRow - 1, FVtblSCol + j, StartRow - 1, FVtblSCol + j, FVtblTitle(j))
    Next j

    ' USDM�̊e�s����͂��Ȃ��珈�����Ă����{�̕���
    For i = StartRow To MaxRow
        classification = "���m��"
        rowsItem = 1 ' ���ږ��̍s���̓f�t�H���g��1�Ƃ���
    
        ' USDM�̍s����͂��Ēl���擾
        For j = Level1Col To MaxCol
            colSpan = 1
            If ws.Cells(i, j).Value <> "" Then ' �����L�q����Ă���
                If classification = "���m��" Then ' ���̍��ڂ��L�q����Ă���̂��������Ă���Ȃ���������͕s�v
                    If IsRequirement(ws.Cells(i, j).Value) Then ' �v���̏ꍇ
                        classification = "�v��"
                    ElseIf IsNintei(ws.Cells(i, j).Value) Then ' �F��d�l�̏ꍇ
                        classification = "�F��d�l"
                        If j <> Level1Col Then ' ����������͂P�w�ڂ����ɋ������\���Ƃ���
                            MsgBox ws.Name + "�V�[�g����������" & i & "�s�ڂłP�w�ڈȊO�ɂ͋�����Ȃ��u�v���v�Ɠ����Z���Ƀ`�F�b�N�{�b�N�X��t����u�F��d�l�v�̕\���������܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Exit Function
                        End If
                    End If
                    If classification = "�v��" Or classification = "�F��d�l" Then ' �v���ƔF��d�l�̏ꍇ
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsReason(ws.Cells(i + 1, j + 1).Value) Then ' ���R�̍s�����݂���ꍇ
                            reason = ws.Cells(i + 1, j + 2).Value
                        Else ' ���R�̍s�͑��݂��Ȃ���΂Ȃ�Ȃ�
                            MsgBox ws.Name + "�V�[�g����������USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                            Exit Function
                        End If
                        rowsItem = 2
                        If IsDescription(ws.Cells(i + 2, j + 1).Value) Then ' �����̍s�����݂���ꍇ
                            rowsItem = 3
                            description = ws.Cells(i + 2, j + 2).Value
                        End If
                    ElseIf IsReason(ws.Cells(i, j).Value) Then ' ���R����n�܂����ꍇ
                        classification = "���R"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsDescription(ws.Cells(i, j).Value) Then ' ��������n�܂����ꍇ
                        classification = "����"
                        content = ws.Cells(i, j + 1).Value
                    ElseIf IsSpec(ws.Cells(i, j).Value) Then ' �Z���̒������̏�񂩂�d�l�Ɣ��f�����ꍇ
                        classification = "�d�l"
                        identifier = ws.Cells(i, j + 1).Value
                        content = ws.Cells(i, j + 2).Value
                        If IsRequirement(identifier) Then ' �d�l�ԍ�������͂��̉E�ׂ̃Z���Ɂu�v���v������ꍇ
                            classification = "�F��d�l" ' �F�������߂��@����͐������[��
                            ' �F��d�l�͌��X�v���ł��邩��A��������̏����́u�v���v�Ɠ����B���������ׂĂ̗�j�̈�E�ɂ���Ă���
                            identifier = ws.Cells(i, j + 1 + 1).Value
                            content = ws.Cells(i, j + 1 + 2).Value
                            If IsReason(ws.Cells(i + 1, j + 1 + 1).Value) Then ' ���R�̍s�����݂���ꍇ
                                reason = ws.Cells(i + 1, j + 1 + 2).Value
                            Else ' ���R�̍s�͑��݂��Ȃ���΂Ȃ�Ȃ�
                                MsgBox ws.Name + "�V�[�g����������USDM�Ő�������Ȃ����R�̍s�������Ȃ��F��d�l(�v��)��" & i & "�s�Ɍ����܂����̂ŁA���̃V�[�g�̈ȍ~�̏����𒆎~���܂��B"
                                Exit Function
                            End If
                            rowsItem = 2
                            If IsDescription(ws.Cells(i + 2, j + 1 + 1).Value) Then ' �����̍s�����݂���ꍇ
                                rowsItem = 3
                                description = ws.Cells(i + 2, j + 1 + 2).Value
                            End If
                        End If
                    ElseIf StrComp(Left(ws.Cells(i, j).Value, 1), "<", vbTextCompare) = 0 And StrComp(Right(ws.Cells(i, j).Value, 1), ">", vbTextCompare) = 0 Then ' <>�ň͂܂�Ă���ꍇ
                        classification = "�d�l�̃O���[�v" ' �Ƃ܂��͉��肷��
                        content = Mid(ws.Cells(i, j).Value, 2, Len(ws.Cells(i, j).Value) - 2) ' <>�̒������o��
                        If StrComp(Left(content, 1), "<", vbTextCompare) = 0 And StrComp(Right(content, 1), ">", vbTextCompare) = 0 Then ' �Ă�<>�ň͂܂�Ă���ꍇ
                            classification = "�d�l�����"
                        ElseIf StrComp(ws.Cells(i + 1, j).Value, "�v��", vbTextCompare) = 0 Then ' �^���̃Z�����v���ł���ꍇ
                            classification = "�v���̃O���[�v"
                        ElseIf IsNintei(ws.Cells(i + 1, j).Value) Then ' �^���̃Z�����F��d�l�ł���ꍇ
                            classification = "�v���̃O���[�v"
                        End If
                    Else ' �ǂݔ�΂��ׂ��ꍇ
                        classification = "�ǂݔ�΂��ׂ�"
                    End If
                Else
                    Exit For ' �������̉�͂���߂�
                End If
            End If
     
            If ws.Cells(i, j).MergeCells Then
                colSpan = ws.Cells(i, j).MergeArea.Columns.Count
            End If
            j = j + colSpan - 1
        Next j
        
        ' �s�̉�͌��ʂɉ�����FV�\�̉��������e���o��
        Select Case classification
            Case "�v��"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Validation", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[���R�]�L�F" & reason & "]" & vbLf & "[�v���]�L�F" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "�Ⴆ�Έ��q��񋓂��܂�"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "�Ⴆ�Αg�����e�X�g, �V�i���I�e�X�g"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            ' Case "���R"
            ' Case "����"
            Case "�d�l"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Verification", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[�d�l�]�L�F" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "�Ⴆ�Έ��q��񋓂��܂�"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "�Ⴆ�΃f�V�W�����e�[�u��"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            Case "�F��d�l"
                Call MergeForce(i, FVtblSCol, i + rowsItem - 1, FVtblSCol, "Validation", "Validation,Verification")
                Call MergeForce(i, FVtblSCol + 1, i + rowsItem - 1, FVtblSCol + 1, identifier)
                Call MergeForce(i, FVtblSCol + 2, i + rowsItem - 1, FVtblSCol + 2, "[���R�]�L�F" & reason & "]" & vbLf & "[�v���]�L�F" & content & "]")
                Call MergeForce(i, FVtblSCol + 3, i + rowsItem - 1, FVtblSCol + 3, "") ' "�Ⴆ�Έ��q��񋓂��܂�"
                Call MergeForce(i, FVtblSCol + 4, i + rowsItem - 1, FVtblSCol + 4, "") ' "�Ⴆ�Αg�����e�X�g"
                Call MergeForce(i, FVtblSCol + 5, i + rowsItem - 1, FVtblSCol + 5, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 6, i + rowsItem - 1, FVtblSCol + 6, "���]��", "���]��,��,��,��")
                Call MergeForce(i, FVtblSCol + 7, i + rowsItem - 1, FVtblSCol + 7, "")
            ' Case "�v���̃O���[�v"
            ' Case "�d�l�̃O���[�v"
            ' Case "�O���[�v�����"
            ' Case Else
        End Select
    Next i
    
    MsgBox "FV�\�̑}���͐���ɍŌ�܂ŏ�������܂����B"
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

