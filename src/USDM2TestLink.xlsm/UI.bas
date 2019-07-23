Attribute VB_Name = "UI"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' XML�錾�́A�}�C�N���\�t�g�̃��C�u�����������R�[�h�ɂ��Đ����I�Ȃ��̂��o�͂��Ȃ��̂ŁA�����ŗp�ӂ���������𒼐ڃX�g���[���ɏo�͂���
Const XMLDeclaration As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbLf

Const cautionStr As String = "�{�����͓��͕����t�H���_���̃t�@�C�����̂��̂�ύX���Ă��܂��܂��B" & vbLf & _
    "�iUSDM�ƔF�������V�[�g��FV�\�̗��}�����܂��B�j" & vbLf & _
    "������ɔ����āA���̓t�H���_�S�̂̃o�b�N�A�b�v���Ƃ��Ă�����s���Ă��������B" & vbLf & _
    "�o�b�N�A�b�v�������̏ꍇ�́A�u�L�����Z���v�������ď����𒆎~���Ă��������B"

Sub startFull()
    start "XML�o�͏���"
End Sub

Sub startFVtblOnly()
    start "FV�\�����̂�"
End Sub

' �ŏ��ɂ�����Ăяo���Ηǂ��̂����A��L2�̃��[�h������
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
    
    rc = MsgBox(cautionStr, vbOKCancel, "�x���I")
    If Not rc = vbOK Then
        MsgBox "�����𒆎~���܂���"
        Exit Sub
    End If
    
    Worksheets("�k�������ʋL�^�l").Activate
    Set ws = ActiveSheet
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    Call ClearList
    srcPath = Worksheets("�kXML�o�͎w�����ݒ�l").Range("���̓p�X").Value
    destPath = Worksheets("�kXML�o�͎w�����ݒ�l").Range("�o�̓p�X").Value
    
    If srcPath = "" Or Dir(srcPath, vbDirectory) = "" Then ' ���̓p�X���w�肳��Ă��Ȃ������݂��Ȃ��Ƃ�
        srcPath = getPath("�����Ώۃt�@�C�����i�[����Ă���t�H���_��I��")
    End If
    If srcPath = "" Then
        MsgBox "���݂�����̓p�X���m�肳��Ȃ������̂ŏ������I�����܂��B"
        Exit Sub
    End If
    
    If mode = "XML�o�͏���" Then
        If destPath = "" Or Dir(destPath, vbDirectory) = "" Then ' �o�̓p�X���w�肳��Ă��Ȃ������݂��Ȃ��Ƃ�
            destPath = getPath("�����ɂ�萶�������t�@�C�����i�[�����t�H���_��I��")
        End If
        If destPath = "" Then
            MsgBox "���݂���o�̓p�X���m�肳��Ȃ������̂ŏ������I�����܂��B"
            Exit Sub
        End If
        reqPath = ExportPath(destPath, "�v��")
        testPath = ExportPath(destPath, "�e�X�g")
        logPath = ExportPath(destPath, "���O")
        
        If reqPath = "" Then
            MsgBox "�v�������o�͂���p�X���m�ۂł��Ȃ������̂ŏ������I�����܂��B"
            Exit Sub
        End If
        If testPath = "" Then
            MsgBox "�e�X�g�����o�͂���p�X���m�ۂł��Ȃ������̂ŏ������I�����܂��B"
            Exit Sub
        End If
        If logPath = "" Then
            MsgBox "���O�����o�͂���p�X���m�ۂł��Ȃ������̂ŏ������I�����܂��B"
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
            If Not ws.ListObjects("�����L�^�e�[�u��").Range.Find(destBaseName, LookAt:=xlWhole) Is Nothing Then ' ���ɓ����V�[�g�����g���Ă��Ȃ����H
                destBaseName = destBaseName & "(" & srcBook.Name & ")"
            End If
            If recognizeUSDMStructure(srcSheet, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) Then
                formType = "USDM�ƔF��"
            Else
                formType = "USDM�ł͂Ȃ�"
            End If
            Set newRow = ws.ListObjects("�����L�^�e�[�u��").ListRows.Add
            If formType = "USDM�ƔF��" Then
                fvTableStatus = "�������s" ' ������
                resultMsg = "�������s" ' ������
                If FVtblSCol = 0 Then ' FV�\�̏�Ԃ��m�F����
                    If InsertFVtbl(srcSheet, MaxRow, MaxCol, StartRow, Level1Col, CategoryCol, RemarksCol, FVtblSCol) And FVtblSCol > 0 Then
                        fvTableStatus = "���񐶐�"
                        srcBook.Save
                    End If
                Else
                    fvTableStatus = "����"
                End If
                If FVtblSCol > 0 And mode = "XML�o�͏���" Then
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
                    resultMsg = "�����ς�"
                    newRow.Range = Array(fileName, srcSheet.Name, formType, fvTableStatus, resultMsg, reqFileName, testFileName, srcPath, reqPath, testPath, logPath)
                Else
                    newRow.Range = Array(fileName, srcSheet.Name, formType, fvTableStatus, "�\", "�\", "�\", srcPath, "�\", "�\", "�\")
                End If
            Else
                newRow.Range = Array(fileName, srcSheet.Name, formType, "�\", "�\", "�\", "�\", srcPath, "�\", "�\", "�\")
            End If
        Next i
        srcBook.Close False
        fileName = Dir()
    Loop
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "�������ɃG���[���������܂����B"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub getInputPath()
    Worksheets("�kXML�o�͎w�����ݒ�l").Range("���̓p�X").Value = getPath("�����Ώۃt�@�C�����i�[����Ă���t�H���_��I��")
End Sub

Sub getOutputPath()
    Worksheets("�kXML�o�͎w�����ݒ�l").Range("�o�̓p�X").Value = getPath("�����ɂ�萶�������t�@�C�����i�[�����t�H���_��I��")
End Sub

' ���[�U�Ƀt�H���_��I�������āA���̃p�X�𓾂�
Function getPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    If fileDialog.Show = -1 Then
        getPath = fileDialog.SelectedItems(1)
    End If
End Function

' �k�������ʋL�^�l�V�[�g�̑S�f�[�^���N���A
Sub ClearList()
    Dim ws As Worksheet
    Worksheets("�k�������ʋL�^�l").Activate
    Set ws = ActiveSheet
    If ActiveSheet.FilterMode Then
        ws.ShowAllData
    End If
    If Not ws.ListObjects("�����L�^�e�[�u��").DataBodyRange Is Nothing Then
        ws.ListObjects("�����L�^�e�[�u��").DataBodyRange.ClearContents
    End If
    ws.ListObjects("�����L�^�e�[�u��").Resize Range("A1:K2")
End Sub

' �o�͂���V�K��Path���m�ۂ���
' basePath�̉���createFolderName�̃t�H���_��V�K�ɍ쐬����B
' ���ꂪ�����̏ꍇ��(1)���珇�ɑ��݂��Ȃ�(N)�܂ł�t�^�������O�𐶐�����
' �K���V����(�������)�t�H���_�𐶐����Ă��̃p�X��Ԃ�
Function ExportPath(basePath As String, createFolderName As String) As String
    ExportPath = ""
    Dim DirectoryExist, DirectoryPath As String
    Dim i As Long

    ' �w��̊����o�̓t�H���_�̒��Ɏw��̖��O�̃t�H���_�[�����
    If StrComp(Right(basePath, 1), "\", vbTextCompare) <> 0 Then
        basePath = basePath & "\"
    End If

    DirectoryPath = basePath & createFolderName
    DirectoryExist = Dir(DirectoryPath, vbDirectory)

    If DirectoryExist = "" Then
        MkDir DirectoryPath
        ExportPath = DirectoryPath
    Else ' �����̖��O�Ƃ����Ԃ����犇�ʂƔԍ���t���ĐV�������O������
        For i = 1 To 1000 ' ���ۂɂ͂���Ȃɐ���������Ǘ��ł��Ȃ����낤
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

' XML�t�@�C���̏o��
Function OutputFile(XML As MSXML2.DOMDocument60, fileName As String) As Boolean
    OutputFile = False
    Dim Reader As New SAXXMLReader60
    Dim writer As New MXXMLWriter60

    writer.indent = True
    writer.standalone = True
    
    ' ============ �}�C�N���\�t�g���C�u�����̕s���������� ============
    'writer.Encoding = "UTF-8"
    writer.Encoding = "shift_jis"
    writer.omitXMLDeclaration = True ' XML�錾�͕����R�[�h�ɂ��Đ����I�Ȃ��̂��o�͂���Ȃ��̂ŁA�p�ӂ���������𒼐ڃX�g���[���ɏo�͂���
    ' ======== �}�C�N���\�t�g���C�u�����̕s����������@����� ========
    
    Set Reader.contentHandler = writer
    Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)

    'ADODB.Stream�I�u�W�F�N�g�𐶐�
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8�ւ̕ϊ��̂��߂Ɏg��
    
    Reader.Parse XML.XML
    
    ' ======== TestLink���̃o�O��������邽�߂̏��� ========
    Dim str As String
    ' TestLink���̃o�O�Ǝv���邪�A<status><CDATA>�^�O�Ԃɉ��s��󔒂�����ƃG���[�ɂȂ�̂ł�����������
    str = Replace(writer.output, "<status>��![CDATA[D]]��</status>", "<status><![CDATA[D]]></status>")
    ' TestLink���̃o�O�Ǝv���邪�A<type><CDATA>�^�O�Ԃɉ��s��󔒂�����ƃG���[�ɂȂ�̂ł�����������
    str = Replace(str, "<type>��![CDATA[3]]��</type>", "<type><![CDATA[3]]></type>")
    ' ===== TestLink���̃o�O��������邽�߂̏�������� =====
    
    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText XMLDeclaration ' �}�C�N���\�t�g���C�u�����̕s����
        .LineSeparator = adCRLF
'        .WriteText Replace(writer.output, vbCrLf, vbLf)
        .WriteText Replace(str, vbCrLf, vbLf) ' TestLink���̃o�O���
        .LineSeparator = adLF
        ' BOM���폜����
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

' �kXML�o�͎w�����ݒ�l�V�[�g�̐ݒ�l�̓ǂݎ��
Function GetSetValues() As Boolean
    GetSetValues = False
    Dim i As Long
        
    separateReqConf = False ' �v���̗��R������̎戵�͎d�l�̒��Ɋ܂߂�����f�t�H���g
    If Worksheets("�kXML�o�͎w�����ݒ�l").Range("�v���̗��R������̎戵").Value = "���ꂼ��J�X�^���t�B�[���h�ɐU�蕪����" Then
        separateReqConf = True
    End If
    
    separateSpcConf = False ' �d�l�̗��R������̎戵�͎d�l�̒��Ɋ܂߂�����f�t�H���g
    If Worksheets("�kXML�o�͎w�����ݒ�l").Range("�d�l�̗��R������̎戵").Value = "���ꂼ��J�X�^���t�B�[���h�ɐU�蕪����" Then
        separateSpcConf = True
    End If
    
    categoryOutConf = True ' �J�e�S�����͏o�͂���̂��f�t�H���g
    sheetNameFirstConf = False ' �J�e�S�����ɂ̓Z�����ɂ�������g���̂��f�t�H���g
    If Worksheets("�kXML�o�͎w�����ݒ�l").Range("�J�e�S���[�̎戵").Value = "�g��Ȃ�(�o�͂��Ȃ�)" Then
        categoryOutConf = False
    ElseIf Worksheets("�kXML�o�͎w�����ݒ�l").Range("�J�e�S���[�̎戵").Value = "�V�[�g�����J�e�S���[�Ƃ��Ďg�p����" Then
        sheetNameFirstConf = True
    End If
    
    If Not Worksheets("�kXML�o�͎w�����ݒ�l").Range("�O���[�v���o�͂̈���").Value = "�O���[�v�������̒P�ƃm�[�h�͏o�͂��Ȃ�" Then
        MsgBox "�ݒ荀�ڂ́u�O���[�v�v�ł����A���݂́u�O���[�v�������̒P�ƃm�[�h�͏o�͂��Ȃ��v������������Ă��܂���B"
        Exit Function
    End If
    
    remarksOutConf = True ' ���l���̏��͏o�͂���̂��f�t�H���g
    If Worksheets("�kXML�o�͎w�����ݒ�l").Range("���l���̎戵").Value = "�o�͂��Ȃ�" Then
        remarksOutConf = False
        MsgBox "���l���̏����o�͂��Ȃ����Ƃɂ���"
    End If
    
    separateFVConf = False ' FV�\�̖ړI�@�\�̎戵�͖ړI�@�\�����ؓ��e�Ɋ܂߂�����f�t�H���g
    If Worksheets("�kXML�o�͎w�����ݒ�l").Range("FV�\�̖ړI�@�\�̎戵").Value = "�ړI�@�\���J�X�^���t�B�[���h�ɐU�蕪����" Then
        separateFVConf = True
    End If

    MaxCheckBoxConf = Worksheets("�kXML�o�͎w�����ݒ�l").Range("�`�F�b�N�{�b�N�X��").Value
    For i = 1 To MaxCheckBoxConf
        CheckBoxSemConf(i) = RemoveSpaces(removeCRLF(Worksheets("�kXML�o�͎w�����ݒ�l").Range("�`�F�b�N�{�b�N�X�̈Ӗ�")(i).Value))
        If CheckBoxSemConf(i) = "" Then ' �\�̕����񂪋󂾂�����f�t�H���g�l�ɂ���
            CheckBoxSemConf(i) = "�`�F�b�N�{�b�N�X" + CStr(i)
        End If
    Next i
    
    GetSetValues = True
End Function

