Attribute VB_Name = "XMLDOM"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' �v���܂��͔F��d�l��v�fReqSpecElement�Ƃ��Ēǉ�
Function appendReqElement(XML, CurrentParentElement As IXMLDOMElement, targetElement As IXMLDOMElement, _
    identifier As String, RorS As String, checkBoxes As String, content As String, reason As String, description As String, _
    currentLevel, currentReqGrp As String, currentCat As String, remarks As String, remarksColFlg, _
    order, ByVal ws As Worksheet, row As Long, fvcol As Long, Optional VVchoice As String = "") As Boolean
    appendReqElement = False
    Dim ReqSpecElement As IXMLDOMElement
    Dim TypeElement As IXMLDOMElement
    Dim NodeOrderElement As IXMLDOMElement
    Dim ContentElement   As IXMLDOMElement
    Dim CustomFieldsElement As IXMLDOMElement
    Dim CustomFieldElement As IXMLDOMElement
    Dim NameElement As IXMLDOMElement
    Dim ValueElement As IXMLDOMElement
    Dim CommentElement As IXMLDOMComment
    
    Set ReqSpecElement = XML.createElement("req_spec")
    Set targetElement = ReqSpecElement
    Call targetElement.setAttribute("doc_id", IDprefix & identifier)
    Call targetElement.setAttribute("title", makeTitle(content))
    Set TypeElement = XML.createElement("type")
    Call TypeElement.appendChild(XML.createCDATASection(System))

    Set NodeOrderElement = XML.createElement("node_order")
    Call NodeOrderElement.appendChild(XML.createCDATASection(order))
    Set ContentElement = XML.createElement("scope")
    If separateReqConf Then
        Call ContentElement.appendChild(XML.createCDATASection(toHTML(content)))
    Else
        Call ContentElement.appendChild(XML.createCDATASection(concatToHTML(content, reason, description)))
    End If
    Set CustomFieldsElement = XML.createElement("custom_fields")
    
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�v���d�l�敪"))
    Call ValueElement.appendChild(XML.createCDATASection(RorS))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    If separateReqConf Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("���R"))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(reason)))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("����"))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(description)))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If

    If StrComp(RorS, "�F��d�l") = 0 Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Set CommentElement = XML.createComment(checkBoxes)
        Call NameElement.appendChild(XML.createCDATASection("�d�l�`�F�b�N�{�b�N�X"))
        Call ValueElement.appendChild(XML.createCDATASection(CheckBox2ValStr(checkBoxes)))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call CustomFieldsElement.appendChild(CommentElement)
    End If
    
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�O���[�v��"))
    Call ValueElement.appendChild(XML.createCDATASection(currentReqGrp))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
    
    If categoryOutConf Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�J�e�S���[��"))
        Call ValueElement.appendChild(XML.createCDATASection(currentCat))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If
    If remarksOutConf And remarksColFlg Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("���l"))
        Call ValueElement.appendChild(XML.createCDATASection(remarks))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If
    
    Call targetElement.appendChild(TypeElement)
    Call targetElement.appendChild(NodeOrderElement)
    Call targetElement.appendChild(ContentElement)
    Call targetElement.appendChild(CustomFieldsElement)
    If StrComp(RorS, "�v��") = 0 Then
        Call appendFVtableElement(XML, targetElement, identifier, makeFVtblTitle(content), order, ws, row, fvcol)
    Else ' �F��d�l�̎�
        Call appendFVtableElement(XML, targetElement, identifier, makeFVtblTitle(content), order, ws, row, fvcol, VALIDATION)
        Call appendFVtableElement(XML, targetElement, identifier + IDsuffix, makeFVtblTitle(content), order, ws, row, fvcol, VERIFICATION) ' �F��d�l�̎��Ɏ����I��Verification�̍��ڂ�t��������
    End If
    Call CurrentParentElement.appendChild(targetElement)
    appendReqElement = True
End Function

' �d�l��v�fReqSpecElement�Ƃ��Ēǉ�
Function appendSpecElement(XML, targetElement As IXMLDOMElement, _
    identifier As String, RorS As String, checkBoxes As String, content As String, currentSpcGrp As String, currentCat As String, _
    remarks As String, remarksColFlg, _
    reqOrder, spcOrder, ByVal ws As Worksheet, row As Long, fvcol As Long) As Boolean
    appendSpecElement = False
    Dim strSpec(1 To 3) As String
    Dim ReqSpecElement As IXMLDOMElement
    Dim TypeElement As IXMLDOMElement
    Dim NodeOrderElement As IXMLDOMElement
    Dim ContentElement   As IXMLDOMElement
    Dim CustomFieldsElement As IXMLDOMElement
    Dim CustomFieldElement As IXMLDOMElement
    Dim NameElement As IXMLDOMElement
    Dim ValueElement As IXMLDOMElement
    Dim CommentElement As IXMLDOMComment
    
    Set ReqSpecElement = XML.createElement("req_spec")
    Call ReqSpecElement.setAttribute("doc_id", identifier)
    Call ReqSpecElement.setAttribute("title", makeTitle(content))
    
    Set TypeElement = XML.createElement("type")
    Call TypeElement.appendChild(XML.createCDATASection(System))
    Set NodeOrderElement = XML.createElement("node_order")
    Call NodeOrderElement.appendChild(XML.createCDATASection(spcOrder))
    Set ContentElement = XML.createElement("scope")
    If separateSpcConf Then
        Call separateSpec(content, strSpec)
        Call ContentElement.appendChild(XML.createCDATASection(toHTML(strSpec(1))))
    Else
        Call ContentElement.appendChild(XML.createCDATASection(toHTML(content)))
    End If
    Set CustomFieldsElement = XML.createElement("custom_fields")
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�v���d�l�敪"))
    Call ValueElement.appendChild(XML.createCDATASection("�d�l"))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
    If separateSpcConf Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("���R"))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(strSpec(2))))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("����"))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(strSpec(3))))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If
    
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Set CommentElement = XML.createComment(checkBoxes)
    Call NameElement.appendChild(XML.createCDATASection("�d�l�`�F�b�N�{�b�N�X"))
    Call ValueElement.appendChild(XML.createCDATASection(CheckBox2ValStr(checkBoxes)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
    Call CustomFieldsElement.appendChild(CommentElement)

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�O���[�v��"))
    Call ValueElement.appendChild(XML.createCDATASection(currentSpcGrp))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    If categoryOutConf Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�J�e�S���[��"))
        Call ValueElement.appendChild(XML.createCDATASection(currentCat))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If

    If remarksOutConf And remarksColFlg Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("���l"))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(remarks)))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If

    Call ReqSpecElement.appendChild(TypeElement)
    Call ReqSpecElement.appendChild(NodeOrderElement)
    Call ReqSpecElement.appendChild(ContentElement)
    Call ReqSpecElement.appendChild(CustomFieldsElement)
    
    Call appendFVtableElement(XML, ReqSpecElement, identifier, makeFVtblTitle(content), reqOrder, ws, row, fvcol)
    
    Call targetElement.appendChild(ReqSpecElement)
                
    appendSpecElement = True
End Function

' �e�X�g�v���iFV�\�j��v�fRequirementElement�Ƃ��Ēǉ�
Function appendFVtableElement(XML, ReqSpecElement As IXMLDOMElement, identifier, _
    title, order, ByVal ws As Worksheet, row As Long, fvcol As Long, Optional VVchoice As String = "") As Boolean
    appendFVtableElement = False
    Dim RequirementElement   As IXMLDOMElement
    Dim DocIdElement As IXMLDOMElement
    Dim TitleElement As IXMLDOMElement
    Dim VersionElement As IXMLDOMElement
    Dim RevisionElement As IXMLDOMElement
    Dim NodeOrderElement As IXMLDOMElement
    Dim DescriptionElement As IXMLDOMElement
    Dim StatusElement As IXMLDOMElement
    Dim TypeElement As IXMLDOMElement
    Dim ExpectedCoverageElement As IXMLDOMElement
    Dim CustomFieldsElement As IXMLDOMElement
    Dim CustomFieldElement As IXMLDOMElement
    Dim NameElement As IXMLDOMElement
    Dim ValueElement As IXMLDOMElement
    Dim vContent As String
    Dim fContent As String
    
    fContent = ws.Cells(row, fvcol + 2).Value
    If StrComp(VVchoice, VERIFICATION, vbTextCompare) = 0 Then ' �F��d�l�̎��Ɏ����I��Verification�̍��ڂ�t��������
        vContent = "�k�F��d�l�ɂ��āA�Ó����m�F�����ł͂Ȃ����؂����{���邱�Ƃ𐄏����܂��B�l" & ws.Cells(row, fvcol + 3).Value
    ElseIf StrComp(VVchoice, VALIDATION, vbTextCompare) = 0 Then ' �F��d�l�̎��ɂ͋����I��Validation(Verification�̍��ڂ͕ʂɕt��������̂Łj
        vContent = ws.Cells(row, fvcol + 3).Value
    Else
        vContent = ws.Cells(row, fvcol + 3).Value
        VVchoice = ws.Cells(row, fvcol).Value
    End If
    
    Set RequirementElement = XML.createElement("requirement")
    Set DocIdElement = XML.createElement("docid")
    Call DocIdElement.appendChild(XML.createCDATASection(IDprefix & identifier & FVsuffix))
    Set TitleElement = XML.createElement("title")
    Call TitleElement.appendChild(XML.createCDATASection(title))
    Set VersionElement = XML.createElement("version")
    Call VersionElement.appendChild(XML.createTextNode("1"))
    Set RevisionElement = XML.createElement("revision")
    Call RevisionElement.appendChild(XML.createTextNode("1"))
    Set NodeOrderElement = XML.createElement("node_order")
    Call NodeOrderElement.appendChild(XML.createTextNode("1")) ' �K����������������Ȃ��̂łP�ɌŒ�
    Set DescriptionElement = XML.createElement("description")
    If separateFVConf Then
        Call DescriptionElement.appendChild(XML.createCDATASection(toHTML(vContent)))
    Else
        Call DescriptionElement.appendChild(XML.createCDATASection(concatFVToHTML(fContent, vContent)))
    End If
    Set StatusElement = XML.createElement("status")
'    Call StatusElement.appendChild(xml.createCDATASection(ReqSTATUS))
    Call StatusElement.appendChild(XML.createTextNode("��![CDATA[" & ReqSTATUS & "]]��")) ' TestLink���̃o�O�Ǝv���邪�A<status><CDATA>�^�O�Ԃɉ��s��󔒂�����ƃG���[�ɂȂ�̂ł�����������
    Set TypeElement = XML.createElement("type")
'    Call TypeElement.appendChild(xml.createCDATASection(ReqTYPE))
    Call TypeElement.appendChild(XML.createTextNode("��![CDATA[" & ReqTYPE & "]]��")) ' TestLink���̃o�O�Ǝv���邪�A<type><CDATA>�^�O�Ԃɉ��s��󔒂�����ƃG���[�ɂȂ�̂ł�����������
    Set ExpectedCoverageElement = XML.createElement("expected_coverage")
    Call ExpectedCoverageElement.appendChild(XML.createCDATASection("1"))
    Set CustomFieldsElement = XML.createElement("custom_fields")
        
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(VVItem))
    Call ValueElement.appendChild(XML.createCDATASection(VVchoice))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(TBidItem))
    Call ValueElement.appendChild(XML.createCDATASection(toCDATA(ws.Cells(row, fvcol + 1).Value)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    If separateFVConf Then
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection(FItem))
        Call ValueElement.appendChild(XML.createCDATASection(toCDATA(fContent)))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
    End If

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(TItem))
    Call ValueElement.appendChild(XML.createCDATASection(toCDATA(ws.Cells(row, fvcol + 4).Value)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(PdRItem))
    Call ValueElement.appendChild(XML.createCDATASection(toCDATA(ws.Cells(row, fvcol + 5).Value)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(PjRItem))
    Call ValueElement.appendChild(XML.createCDATASection(toCDATA(ws.Cells(row, fvcol + 6).Value)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)

    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection(FLFPItem))
    Call ValueElement.appendChild(XML.createCDATASection(toCDATA(ws.Cells(row, fvcol + 7).Value)))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
        
    Call RequirementElement.appendChild(DocIdElement)
    Call RequirementElement.appendChild(TitleElement)
    Call RequirementElement.appendChild(VersionElement)
    Call RequirementElement.appendChild(RevisionElement)
    Call RequirementElement.appendChild(NodeOrderElement)
    Call RequirementElement.appendChild(DescriptionElement)
    Call RequirementElement.appendChild(StatusElement)
    Call RequirementElement.appendChild(TypeElement)
    Call RequirementElement.appendChild(ExpectedCoverageElement)
    Call RequirementElement.appendChild(CustomFieldsElement)
        
    Call ReqSpecElement.appendChild(RequirementElement)
    appendFVtableElement = True
End Function

' �e�X�g�X�C�[�g���e�X�g�݌v�J�n�ɓs���̗ǂ��悤�ɏ���������
' �v�fTestSuiteElement�Ƃ��Ēǉ� �i���̂܂܂̍\���𐄏�����킯�ł͂Ȃ��j
Function InitTestSuite(XML, rootElement As IXMLDOMElement) As Boolean
    InitTestSuite = False
    Dim TestSuiteElement As IXMLDOMElement
    Dim NodeOrderElement As IXMLDOMElement
    Dim DetailsElement As IXMLDOMElement
    Dim CustomFieldsElement As IXMLDOMElement
    Dim CustomFieldElement As IXMLDOMElement
    Dim NameElement As IXMLDOMElement
    Dim ValueElement As IXMLDOMElement
    Dim CommentElement As IXMLDOMComment

    Set NodeOrderElement = XML.createElement("node_order")
    Call NodeOrderElement.appendChild(XML.createCDATASection(""))
    Set DetailsElement = XML.createElement("details")
    Call DetailsElement.appendChild(XML.createCDATASection(""))
    
    Set CustomFieldsElement = XML.createElement("custom_fields")
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
    Call ValueElement.appendChild(XML.createCDATASection(""))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
    Call rootElement.appendChild(NodeOrderElement)
    Call rootElement.appendChild(DetailsElement)
    Call rootElement.appendChild(CustomFieldsElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�����g����(HAYST�@)�e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("1"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(haystsample))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection("HAYST-01"))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
        Call AppendSampleTestCase(XML, TestSuiteElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�P�@�\�e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("3"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�L���g�����e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("4"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�֑��g�����e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("5"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "��ԑJ�ڃe�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("6"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�V�i���I�e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("7"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "���׃e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("8"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�Z�L�����e�B�e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("9"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "��@�\�e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("10"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    Set TestSuiteElement = XML.createElement("testsuite")
    Call TestSuiteElement.setAttribute("id", "")
    Call TestSuiteElement.setAttribute("name", "�G���[�����e�X�g")
        Set NodeOrderElement = XML.createElement("node_order")
        Call NodeOrderElement.appendChild(XML.createCDATASection("11"))
        Set DetailsElement = XML.createElement("details")
        Call DetailsElement.appendChild(XML.createCDATASection(""))
        
        Set CustomFieldsElement = XML.createElement("custom_fields")
        Set CustomFieldElement = XML.createElement("custom_field")
        Set NameElement = XML.createElement("name")
        Set ValueElement = XML.createElement("value")
        Call NameElement.appendChild(XML.createCDATASection("�e�X�g�݌v�敪�ԍ�"))
        Call ValueElement.appendChild(XML.createCDATASection(""))
        Call CustomFieldElement.appendChild(NameElement)
        Call CustomFieldElement.appendChild(ValueElement)
        Call CustomFieldsElement.appendChild(CustomFieldElement)
        Call TestSuiteElement.appendChild(NodeOrderElement)
        Call TestSuiteElement.appendChild(DetailsElement)
        Call TestSuiteElement.appendChild(CustomFieldsElement)
    Call rootElement.appendChild(TestSuiteElement)

    InitTestSuite = True
End Function

' �e�X�g�v��(�v������юd�l)��1��1�ɑΉ�����悤�Ƀe�X�g�P�[�X�̗v�f��ǉ����邱�Ƃ�
' �e�X�g�݌v�J�n�ɓs���̗ǂ��悤�ɏ���������B�v�fTestCaseElement�Ƃ��Ēǉ�����B
' �i���̂܂܂̍\���𐄏�����킯�ł͂Ȃ��B���ƂŃe�X�g�P�[�X��ǉ����A�v�����փ����N�𒣂�͎̂�Ԃ�������B
' �@�����Ŏ����������ꂽ���R�[�h���ړ������蕡�������肵�Đ݌v���Ă��������y�ł���B�j
Function appendTestCaseElement(XML, TestSuiteRootElement As IXMLDOMElement, baseIds As String, testSuiteName As String) As Boolean
    appendTestCaseElement = False
    Dim TestCaseElement As IXMLDOMElement
    Dim RequirementsElement As IXMLDOMElement
    Dim RequirementElement As IXMLDOMElement
    Dim DocIdElement As IXMLDOMElement
    Dim CommentElement As IXMLDOMComment
    Dim baseId As String
    Dim xpath As String
    
    Select Case testSuiteName
    Case "�����g����(HAYST�@)�e�X�g"
        xpath = "/testsuite/testsuite[1]"
    Case "�P�@�\�e�X�g"
        xpath = "/testsuite/testsuite[2]"
    Case Else
        MsgBox "�G���[�F���݂��Ȃ��e�X�g�X�C�[�g�����w�肳�ꂽ"
    End Select
    
    If baseIds <> "" Then ' ���̃e�X�g�ɓ�������Ă��܂��Ă���ꍇ�̓R�����g����菜����""�ɂȂ��Ă���
    
        Set TestCaseElement = XML.createElement("testcase")
        Call TestCaseElement.setAttribute("name", baseIds & "�̃e�X�g")
        Set RequirementsElement = XML.createElement("requirements")
    
        baseId = extractId(baseIds)
        Do While baseId <> vbNullString
            Set RequirementElement = XML.createElement("requirement")
            Set DocIdElement = XML.createElement("doc_id")
            Call DocIdElement.appendChild(XML.createCDATASection(baseId))
            Call RequirementElement.appendChild(DocIdElement)
            Call RequirementsElement.appendChild(RequirementElement)
            
            baseId = extractId(baseIds)
        Loop
                    
        Call TestCaseElement.appendChild(RequirementsElement)
        Call TestSuiteRootElement.SelectSingleNode(xpath).appendChild(TestCaseElement)
    End If
    appendTestCaseElement = True
End Function

' �T���v���̃e�X�g�P�[�X��v�fTestCaseElement�Ƃ��Ēǉ�����B(��������̂�������ɐ���������̂ł͂Ȃ�)
Function AppendSampleTestCase(XML, TestSuiteElement As IXMLDOMElement) As Boolean
    AppendSampleTestCase = False
    Dim TestCaseElement As IXMLDOMElement
    Dim NodeOrderElement As IXMLDOMElement
    Dim ExternalIdElement As IXMLDOMElement
    Dim VersionElement As IXMLDOMElement
    Dim SummaryElement As IXMLDOMElement
    Dim PreconditionsElement As IXMLDOMElement
    Dim ExecutionTypeElement As IXMLDOMElement
    Dim ImportanceElement As IXMLDOMElement
    Dim EstimatedExecDurationElement As IXMLDOMElement
    Dim StatusElement As IXMLDOMElement
    Dim IsOpenElement As IXMLDOMElement
    Dim ActiveElement As IXMLDOMElement
    Dim CustomFieldsElement As IXMLDOMElement
    Dim CustomFieldElement As IXMLDOMElement
    Dim NameElement As IXMLDOMElement
    Dim ValueElement As IXMLDOMElement
    Dim CommentElement As IXMLDOMComment

    Set TestCaseElement = XML.createElement("testcase")
    Call TestCaseElement.setAttribute("internalid", "")
    Call TestCaseElement.setAttribute("name", "�g�����e�X�g�̋L�ڌ`���̈��(�Q�ƌ�ɍ폜���Ă�������)")
    Set NodeOrderElement = XML.createElement("node_order")
    Call NodeOrderElement.appendChild(XML.createCDATASection("0"))
    Set ExternalIdElement = XML.createElement("externalid")
    Call ExternalIdElement.appendChild(XML.createCDATASection("4")) ' ����͉���4�Ȃ̂��H
    Set VersionElement = XML.createElement("version")
    Call VersionElement.appendChild(XML.createCDATASection("1"))
    Set SummaryElement = XML.createElement("summary")
    Call SummaryElement.appendChild(XML.createCDATASection(SampleSummary))
    Set PreconditionsElement = XML.createElement("preconditions")
    Call PreconditionsElement.appendChild(XML.createCDATASection(SamplePreconditions))
    Set ExecutionTypeElement = XML.createElement("execution_type")
    Call ExecutionTypeElement.appendChild(XML.createCDATASection("1"))
    Set ImportanceElement = XML.createElement("importance")
    Call ImportanceElement.appendChild(XML.createCDATASection("2"))
    Set EstimatedExecDurationElement = XML.createElement("estimated_exec_duration")
    Call EstimatedExecDurationElement.appendChild(XML.createTextNode(""))
    Set StatusElement = XML.createElement("status")
    Call StatusElement.appendChild(XML.createTextNode("1"))
    Set IsOpenElement = XML.createElement("is_open")
    Call IsOpenElement.appendChild(XML.createTextNode("1"))
    Set ActiveElement = XML.createElement("active")
    Call ActiveElement.appendChild(XML.createTextNode("1"))
    Set CustomFieldsElement = XML.createElement("custom_fields")
    Set CustomFieldElement = XML.createElement("custom_field")
    Set NameElement = XML.createElement("name")
    Set ValueElement = XML.createElement("value")
    Call NameElement.appendChild(XML.createCDATASection("�����e�X�g�݌v�敪�ԍ�"))
    Call ValueElement.appendChild(XML.createCDATASection("HAYST-01"))
    Call CustomFieldElement.appendChild(NameElement)
    Call CustomFieldElement.appendChild(ValueElement)
    Call CustomFieldsElement.appendChild(CustomFieldElement)
    
    Call TestCaseElement.appendChild(NodeOrderElement)
    Call TestCaseElement.appendChild(ExternalIdElement)
    Call TestCaseElement.appendChild(VersionElement)
    Call TestCaseElement.appendChild(SummaryElement)
    Call TestCaseElement.appendChild(PreconditionsElement)
    Call TestCaseElement.appendChild(ExecutionTypeElement)
    Call TestCaseElement.appendChild(ImportanceElement)
    Call TestCaseElement.appendChild(EstimatedExecDurationElement)
    Call TestCaseElement.appendChild(StatusElement)
    Call TestCaseElement.appendChild(IsOpenElement)
    Call TestCaseElement.appendChild(ActiveElement)
    Call TestCaseElement.appendChild(CustomFieldsElement)
    
    Call TestSuiteElement.appendChild(TestCaseElement)

    AppendSampleTestCase = True
End Function

