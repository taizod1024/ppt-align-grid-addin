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
    Dim lcnt As Integer     ' left count
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
               
            If shp.Locked Then
            
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
                lcnt = Round(left / ActivePresentation.GridDistance)
                lrem = left - lcnt * ActivePresentation.GridDistance
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
                        "left   : " + CStr(shp.left) + " -> " + CStr(left) + " " + CStr(lcnt) + " " + CStr(lrem) + vbCrLf + _
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

    Dim linkCnt As Integer
    linkCnt = 0
    ActiveWindow.Selection.Unselect

    ' stackoverflow���Q�l�Ɏ���
    ' https://stackoverflow.com/questions/55724877/how-to-obtain-shapes-to-hyperlinks-in-powerpoint-vba
    
    Dim sld As Slide
    Dim shp As Shape
    Dim actionSetting As actionSetting
    Dim mouseActivation As Variant
    Dim strUrl As String
    Dim strLabel As String
    Dim i As Integer
    Dim j As Integer
       
    ' �X���C�h���
    For Each sld In ActivePresentation.Slides
    
        ' �X���C�h��\��
        ActiveWindow.View.GotoSlide sld.SlideIndex
        
        ' �X���C�h�z���̐}�`���
        For Each shp In sld.Shapes
            
            ' �}�`��I��
            shp.Select
            
            ' �}�`�Ƀe�L�X�g������ꍇ
            If shp.TextFrame.HasText Then
            
                ' ����񋓂��Ă��邩�s��
                For Each mouseActivation In Array(ppMouseClick, ppMouseOver)
                    Set actionSetting = shp.TextFrame.TextRange.ActionSettings(mouseActivation)
                    strUrl = ""
                    
                    ' �e�L�X�g�𕶎����ɗ�
                    For i = 1 To shp.TextFrame.TextRange.Characters.Count
                        Set actionSetting = shp.TextFrame.TextRange.Characters(i).ActionSettings(mouseActivation)
                        
                        ' �����N�̏ꍇ�������N��̕ύX���������ꍇ�̂ݏ���
                        If actionSetting.Action = ppActionHyperlink And strUrl <> actionSetting.Hyperlink.Address Then
                            strUrl = actionSetting.Hyperlink.Address
                            linkCnt = linkCnt + 1
                            
                            ' �����N�؂�̏ꍇ
                            If Not IsExistUrl(strUrl) Then
                            
                                ' �����N�悪�����������W�߂ă��x�����쐬
                                strLabel = ""
                                For j = i To shp.TextFrame.TextRange.Characters.Count
                                    Dim actionSettingLabel As actionSetting
                                    Set actionSettingLabel = shp.TextFrame.TextRange.Characters(j).ActionSettings(mouseActivation)
                                    If actionSettingLabel.Action <> ppActionHyperlink Then Exit For
                                    If strUrl <> actionSettingLabel.Hyperlink.Address Then Exit For
                                    strLabel = strLabel & shp.TextFrame.TextRange.Characters(j).Text
                                Next
                                
                                ' ���x����I��
                                shp.TextFrame.TextRange.Characters(i, Len(strLabel)).Select
                                
                                ' �����I��
                                Exit Sub
                                
                            End If
                        End If
                    Next
                Next
            End If
        Next
    Next
  
    ' �������͑I������
    ActiveWindow.Selection.Unselect
    
    ' �����N�؂��ʒm
    MsgBox CStr(linkCnt) + " ��URL���`�F�b�N���܂����B" + vbCrLf + _
        "�����N�؂��URL�͂���܂���B", vbInformation
    
End Sub

Function IsUrl(strUrl As String) As Boolean

    Const strhttp = "http:"
    Const strhttps = "https:"
    
    IsUrl = left(strUrl, Len(strhttp)) = strhttp Or left(strUrl, Len(strhttps)) = strhttps

End Function

Function IsExistUrl(strUrl As String) As Boolean

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
        IsExistUrl = status <> 404
        Set httpReq = Nothing
        
        Debug.Print "----" + vbCrLf + _
        "url    : " + strUrl + vbCrLf + _
        "status : " + CStr(status) + vbCrLf + _
        "result : " + CStr(IsExistUrl)
        
    Else
    
        Dim path As String
        path = ActivePresentation.path + "\" + strUrl
        IsExistUrl = Dir(path) <> ""
        
        Debug.Print "----" + vbCrLf + _
        "url    : " + strUrl + vbCrLf + _
        "result : " + CStr(IsExistUrl)
        
    End If
        
End Function


