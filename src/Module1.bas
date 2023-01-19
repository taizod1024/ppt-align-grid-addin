Attribute VB_Name = "Module1"
Option Explicit

Sub �O���b�h���ɑ�����_onAction(constrol As IRibbonControl)

    �O���b�h���ɑ�����

End Sub

Sub �Б��ڑ��̃R�l�N�^_onAction(constrol As IRibbonControl)

    �Б��ڑ��̃R�l�N�^

End Sub

Sub �����N�؂��URL_onAction(constrol As IRibbonControl)

    �����N�؂��URL

End Sub

Sub �O���b�h���ɑ�����()

    Dim test As Object          ' test object
    Dim sld As Slide            ' slide
    Dim shprng As ShapeRange    ' shape range
    Dim shp As Shape            ' shape
    Dim shpdic As Object        ' shape dictionary for master
    Dim shpcnt As Integer       ' shape count
    Dim cnnct As Boolean        ' connector or not
    Dim left As Single      ' left
    Dim top As Single       ' top
    Dim width As Single     ' width
    Dim height As Single    ' height
    Dim lCnt As Integer     ' left count
    Dim lrem As Single      ' left remain
    Dim tcnt As Integer     ' top count
    Dim trem As Single      ' top remain
    Dim wcnt As Integer     ' width count
    Dim wrem As Single      ' width remain
    Dim hcnt As Integer     ' height count
    Dim hrem As Single      ' height remain
    
    ' �X���C�h���Ȃ��ꍇ���`�F�b�N
    Set test = Nothing
    On Error Resume Next
    Set test = ActiveWindow.Selection
    On Error GoTo 0
    If test Is Nothing Then
        Debug.Print "----" + vbCrLf + _
            "exit: cannot get selection"
        Exit Sub
    End If
    
    ' �}�`���I�����̓X���C�h�}�X�^�ȊO�̐}�`�����ׂđI��
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
    
        ' �X���C�h�擾
        Set sld = Nothing
        On Error Resume Next
        Set sld = ActivePresentation.Slides.FindBySlideID(ActivePresentation.Windows(1).Selection.SlideRange.SlideID)
        On Error GoTo 0
        If sld Is Nothing Then
            ' �X���C�h�̋��Ԃ�I�����͂����Ŕ�����
            ' �X���C�h�}�X�^�I�����������Ŕ�����
            ' ���{���̓X���C�h�}�X�^���擾�ł���Δ�����K�v�͂Ȃ�
            Debug.Print "----" + vbCrLf + _
                "exit: cannot get slide object"
            Exit Sub
        End If
        
        ' �v���[�X�z���_�̐}�`ID�̎�����
        Set shpdic = CreateObject("Scripting.Dictionary")
        For Each shp In sld.Shapes.Placeholders
            shpdic(shp.Id) = shp.Id
        Next
    
        ' �v���[�X�z���_����уt�b�^�[�̐}�`�ȊO��I��
        shpcnt = 0
        For Each shp In sld.Shapes.Range
        
            If Not shpdic.Exists(shp.Id) And _
                Not shp.Name Like "Footer Placeholder*" And _
                Not shp.Name Like "Slide Number Placeholder*" And _
                Not shp.Name Like "Date Placeholder*" Then
                
                shpcnt = shpcnt + 1
                shp.Select msoFalse
            End If
            
        Next
        
        ' �I�������}�`��0�Ȃ�㑱�̏����ŃG���[�ɂȂ邽�ߏI��
        If shpcnt = 0 Then
            Exit Sub
        End If
                
    End If
    
    ' �I��}�`��Ώۂɏ����J�n
    Set shprng = ActiveWindow.Selection.ShapeRange
    ActiveWindow.Selection.Unselect
    
    ' �I�����Ă���}�`�P�ʂňʒu��������
    shpcnt = 0
    For Each shp In shprng
                       
        ' �R�l�N�^����
        cnnct = False
        If shp.Connector Then
            If shp.ConnectorFormat.BeginConnected Then cnnct = True
            If shp.ConnectorFormat.EndConnected Then cnnct = True
        End If
        
        ' �R�l�N�^�ȊO�̏ꍇ
        If Not cnnct Then
            
            Dim is_locked As Boolean
            is_locked = False
            On Error Resume Next
            is_locked = shp.Locked
            On Error GoTo 0
            
            If is_locked Then
            
                ' ���b�N����Ă���ꍇ
                
                ' �f�o�b�O�F���b�N�̕\��
                Debug.Print "----" + vbCrLf + _
                    "id     : " + CStr(shp.Id) + vbCrLf + _
                    "status : locked"
                        
            Else
                            
                ' ���b�N����Ă��Ȃ��ꍇ
                            
                ' ��Ɨp�̈ʒu���擾
                left = shp.left
                top = shp.top
                width = shp.width
                height = shp.height
                
                ' ���S�ɍ��킹�Ē����A�O���b�h���̑�������Ō��ɖ߂�
                left = left - ActivePresentation.PageSetup.SlideWidth / 2
                top = top - ActivePresentation.PageSetup.SlideHeight / 2
                
                ' �J�Ԑ������߂���ō������v�Z���O���b�h���ɑ����悤����
                lCnt = Round(left / ActivePresentation.GridDistance)
                lrem = left - lCnt * ActivePresentation.GridDistance
                left = left - lrem
                width = width + lrem
                
                tcnt = Round(top / ActivePresentation.GridDistance)
                trem = top - tcnt * ActivePresentation.GridDistance
                top = top - trem
                height = height + trem
                
                wcnt = Round(width / ActivePresentation.GridDistance)
                wrem = width - wcnt * ActivePresentation.GridDistance
                width = width - wrem
                
                hcnt = Round(height / ActivePresentation.GridDistance)
                hrem = height - hcnt * ActivePresentation.GridDistance
                height = height - hrem
                
                ' ���̈ʒu�ɖ߂�
                left = left + ActivePresentation.PageSetup.SlideWidth / 2
                top = top + ActivePresentation.PageSetup.SlideHeight / 2
                
                If Abs(shp.left - left) < 0.01 And _
                    Abs(shp.top - top) < 0.01 And _
                    Abs(shp.width - width) < 0.01 And _
                    Abs(shp.height - height) < 0.01 Then
                    
                    ' �ύX����Ă��Ȃ���Ή������Ȃ�
                Else
                
                    ' �ύX����Ă���ΐV���ɑI������
                    
                    ' �f�o�b�O�F�ύX���e�̕\��
                    Debug.Print "----" + vbCrLf + _
                        "id     : " + CStr(shp.Id) + vbCrLf + _
                        "grid   : " + CStr(ActivePresentation.GridDistance) + vbCrLf + _
                        "left   : " + CStr(shp.left) + " -> " + CStr(left) + " " + CStr(lCnt) + " " + CStr(lrem) + vbCrLf + _
                        "top    : " + CStr(shp.top) + " -> " + CStr(top) + " " + CStr(tcnt) + " " + CStr(trem) + vbCrLf + _
                        "width  : " + CStr(shp.width) + " -> " + CStr(width) + " " + CStr(wcnt) + " " + CStr(wrem) + vbCrLf + _
                        "height : " + CStr(shp.height) + " -> " + CStr(height) + " " + CStr(hcnt) + " " + CStr(hrem) + vbCrLf
    
                    shp.Select msoFalse
                    shpcnt = shpcnt + 1
                    shp.LockAspectRatio = msoFalse
                    shp.left = left
                    shp.top = top
                    shp.width = width
                    shp.height = height
                    
                End If
                
            End If
            
        End If
        
    Next
    
    Exit Sub
    
End Sub
    
Sub �Б��ڑ��̃R�l�N�^()

    Dim sld As Slide    ' slide
    Dim shp As Shape    ' shape
    Dim flg As Boolean  ' flag
    
    ' �X���C�h�擾
    Set sld = Nothing
    On Error Resume Next
    Set sld = ActivePresentation.Slides.FindBySlideID(ActivePresentation.Windows(1).Selection.SlideRange.SlideID)
    On Error GoTo 0
    If sld Is Nothing Then
        ' �X���C�h�̋��Ԃ�I�����͂����Ŕ�����
        ' �X���C�h�}�X�^�I�����������Ŕ�����
        ' ���{���̓X���C�h�}�X�^���擾�ł���Δ�����K�v�͂Ȃ�
        Debug.Print "----" + vbCrLf + _
            "exit: cannot get slide object"
        Exit Sub
    End If
    
    ' �X���C�h�̐}�`�ꗗ
    ActiveWindow.Selection.Unselect
    For Each shp In sld.Shapes.Range
        
        If shp.Connector Then
        
            flg = False
            
            ' �Б��R�l�N�^�̃`�F�b�N
            If shp.ConnectorFormat.BeginConnected And Not shp.ConnectorFormat.EndConnected Then flg = True
            If shp.ConnectorFormat.EndConnected And Not shp.ConnectorFormat.BeginConnected Then flg = True
            
            ' �Б��ڑ��̃R�l�N�^������������I��
            If flg Then
                shp.Select
                Exit Sub
            End If
            
        End If
        
    Next
    
    ' �Б��ڑ��̃R�l�N�^��������Ȃ��������Ƃ�ʒm
    MsgBox "�Б��ڑ��̃R�l�N�^�͂���܂���B", vbInformation
    Exit Sub
    
End Sub

Sub �����N�؂��URL()

    Dim strUrlOld As String
    Dim strUrlNew As String
    Dim lCnt As Integer
    Dim sld As Slide
    Dim shp As Shape
    Dim strMessage As String
    
    On Error GoTo ON_ERROR
    
    ' ����̑I����Ԃ�����
    ActiveWindow.Selection.Unselect
    
    ' �X���C�h��擪�����
    strUrlOld = ""
    strUrlNew = ""
    lCnt = 0
    For Each sld In ActivePresentation.Slides
    
        ' �������̃X���C�h�ֈړ��A.Select�ł͕\���͕ς��Ȃ�
        ActiveWindow.View.GotoSlide sld.SlideIndex
        
        ' �X���C�h�z���̐}�`���
        For Each shp In sld.Shapes
        
            ' �}�`���`�F�b�N
            ' - �`�F�b�N���s����strUrlOld���ɒl���ݒ�
            ' - �`�F�b�N���s��͍Ō�̃X���C�h�܂�URL�̍Đݒ�����{
            CheckShape shp, strUrlOld, strUrlNew, lCnt
        Next
        
    Next
  
    ' �X���C�h�񋓌�
    If lCnt = 0 Then
        ' �`�F�b�N������
        ActiveWindow.Selection.Unselect
        strMessage = "�Ō�܂Ō������܂����B"
        Debug.Print "----" + vbCrLf + strMessage
        MsgBox strMessage, vbInformation
            
    Else
        ' �`�F�b�N���s��
        strMessage = "�����N�؂��URL��������������" + CStr(lCnt) + "���Đݒ肵�܂����B" + vbCrLf + _
            "�ēx���s���Ă��������B"
        Debug.Print "----" + vbCrLf + strMessage
        MsgBox strMessage, vbInformation

    End If
    
    Exit Sub
    
ON_ERROR:

    ' �G���[��
    strMessage = Err.Source
    Debug.Print "----" + vbCrLf + strMessage
    MsgBox strMessage, vbExclamation

End Sub

Sub CheckShape(shp As Shape, _
    ByRef strUrlOld As String, _
    ByRef strUrlNew As String, _
    ByRef lCnt As Integer)

    ' �X�̐}�`���`�F�b�N
    
    ' stackoverflow���Q�l�Ɏ���
    ' https://stackoverflow.com/questions/55724877/how-to-obtain-shapes-to-hyperlinks-in-powerpoint-vba
    
    Dim mouseActivation As Variant
    Dim actionSetting As actionSetting
    Dim rangeLabel As TextRange
    Dim strUrl As String
    Dim strLabel As String
    Dim strType As String
    Dim i As Integer
    Dim j As Integer
    
    ' �}�`��I��
    shp.Select
    
    If shp.Type = msoGroup Then
    
        ' �}�`���O���[�v�̏ꍇ
        
        Dim shp2 As Shape
        For Each shp2 In shp.GroupItems
            CheckShape shp2, strUrlOld, strUrlNew, lCnt
        Next
    
    Else
    
        ' �}�`���O���[�v�łȂ��ꍇ
        
        For Each actionSetting In shp.ActionSettings
        
            ' *** �}�`�̃����N����`�F�b�N ***
            If actionSetting.Action = ppActionHyperlink Then
                Set rangeLabel = Nothing
                strUrl = actionSetting.Hyperlink.Address
                strLabel = ""
                strType = "�}�`"
                CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
            End If
            
        Next
    
        ' *** �����N���ꂽ�}�`�̃����N����`�F�b�N ***
        If shp.Type = msoLinkedPicture Then
            Set rangeLabel = Nothing
            strUrl = shp.LinkFormat.SourceFullName
            strLabel = ""
            strType = "�����N���ꂽ�}�`"
            CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
        ElseIf shp.TextFrame.HasText Then
                    
            For Each mouseActivation In Array(ppMouseClick, ppMouseOver)
                Set actionSetting = shp.TextFrame.TextRange.ActionSettings(mouseActivation)
                
                ' *** �e�L�X�g�S�̂̃����N����`�F�b�N ***
                If actionSetting.Action = ppActionHyperlink Then
                    Set rangeLabel = shp.TextFrame.TextRange.Characters(1, shp.TextFrame.TextRange.Characters.Count)
                    strUrl = actionSetting.Hyperlink.Address
                    strLabel = shp.TextFrame.TextRange.Text
                    strType = "�e�L�X�g"
                    CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
                Else
                                        
                    strUrl = ""
                    For i = 1 To shp.TextFrame.TextRange.Characters.Count
                        Set actionSetting = shp.TextFrame.TextRange.Characters(i).ActionSettings(mouseActivation)
                        
                        ' *** �e�L�X�g�̃����N����`�F�b�N ***
                        If actionSetting.Action = ppActionHyperlink And strUrl <> actionSetting.Hyperlink.Address Then
                            strUrl = actionSetting.Hyperlink.Address
                            strLabel = ""
                            For j = i To shp.TextFrame.TextRange.Characters.Count
                                Dim actionSettingLabel As actionSetting
                                Set actionSettingLabel = shp.TextFrame.TextRange.Characters(j).ActionSettings(mouseActivation)
                                If actionSettingLabel.Action <> ppActionHyperlink Then Exit For
                                If strUrl <> actionSettingLabel.Hyperlink.Address Then Exit For
                                strLabel = strLabel & shp.TextFrame.TextRange.Characters(j).Text
                            Next
                            Set rangeLabel = shp.TextFrame.TextRange.Characters(i, Len(strLabel))
                            Set actionSetting = rangeLabel.ActionSettings(mouseActivation)
                            strType = "�e�L�X�g*"
                            CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
                        End If
                        
                    Next
                End If
            Next
        End If
    End If
  
End Sub

Sub CheckShapeSub(shp As Shape, _
    ByRef strUrlOld As String, _
    ByRef strUrlNew As String, _
    ByRef lCnt As Integer, _
    actionSetting As actionSetting, _
    rangeLabel As TextRange, _
    strUrl As String, _
    strLabel As String, _
    strType As String)

    Dim strStatus As String
    If strUrlNew = "" Then
        strStatus = GetStatus(strUrl)
                
        Debug.Print "----" + vbCrLf + _
                    "type   : " + strType + vbCrLf + _
                    "url    : " + strUrl + vbCrLf + _
                    "status : " + strStatus

        If strStatus <> "" Then
            If Not rangeLabel Is Nothing Then rangeLabel.Select
            
            Do
                strUrlNew = InputBox(strType + "�̐V����URL����͂��Ă��������B" + vbCrLf + strUrl, strStatus, strUrl)
                If strUrlNew = "" Then
                    If MsgBox("����URL���X�L�b�v���đ��s���܂����H" + vbCrLf + "�͂��F�X�L�b�v���đ��s" + vbCrLf + "�������F�I��", vbYesNo + vbQuestion) = vbYes Then
                        Exit Sub
                    Else
                        Err.Raise 1001, "�L�����Z������܂����B"
                    End If
                End If
                strUrl = strUrlNew
                strStatus = GetStatus(strUrl)
                        
                Debug.Print "----" + vbCrLf + _
                            "type   : " + strType + vbCrLf + _
                            "newurl : " + strUrl + vbCrLf + _
                            "status : " + strStatus
                
            Loop While strStatus <> ""
            
            If Not rangeLabel Is Nothing Then If strLabel = strUrl Then rangeLabel.Text = strUrlNew
            If shp.Type = msoLinkedPicture Then shp.LinkFormat.SourceFullName = strUrlNew
            If Not actionSetting Is Nothing Then actionSetting.Hyperlink.Address = strUrlNew
            strUrlOld = strUrl
            lCnt = lCnt + 1
            
        End If
        
    ElseIf strUrl = strUrlOld Then
    
        If Not rangeLabel Is Nothing Then If strLabel = strUrl Then rangeLabel.Text = strUrlNew
        If shp.Type = msoLinkedPicture Then shp.LinkFormat.SourceFullName = strUrlNew
        If Not actionSetting Is Nothing Then actionSetting.Hyperlink.Address = strUrlNew
        lCnt = lCnt + 1
    End If

End Sub

Function IsUrl(strUrl As String) As Boolean

    Const strhttp = "http:"
    Const strhttps = "https:"
    
    IsUrl = left(strUrl, Len(strhttp)) = strhttp Or left(strUrl, Len(strhttps)) = strhttps

End Function

Function GetStatus(strUrl As String) As String

    If IsUrl(strUrl) Then
    
        ' �ȉ��̃T�C�g���Q�l�Ɏ���
        ' https://tonari-it.com/excel-vba-http-request/
    
        Dim httpReq
        Set httpReq = CreateObject("MSXML2.XMLHTTP")
        httpReq.Open "GET", strUrl
        httpReq.Send
        
        Do While httpReq.readyState < 4
            DoEvents
        Loop
        
        Dim status As Integer
        status = httpReq.status
        GetStatus = status
        If status = 200 Then GetStatus = ""
        If status = 404 Then GetStatus = CStr(status) + " NOT FOUND"
        If status = 12007 Then GetStatus = CStr(status) + " ERROR_WINHTTP_NAME_NOT_RESOLVED"
        Set httpReq = Nothing
        
        Debug.Print "----" + vbCrLf + _
        "url    : " + strUrl + vbCrLf + _
        "code   : " + CStr(status) + vbCrLf + _
        "status : " + GetStatus
        
    Else
    
        Dim path As String
        path = ActivePresentation.path + "\" + strUrl
        GetStatus = IIf(Dir(path) <> "", "", "FILE NOT FOUND")
        
    End If
        
End Function


