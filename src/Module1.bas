Attribute VB_Name = "Module1"
Option Explicit

Sub �O���b�h���ɑ�����_onAction(constrol As IRibbonControl)

    �O���b�h���ɑ�����

End Sub

Sub �Б��ڑ��̃R�l�N�^_onAction(constrol As IRibbonControl)

    �Б��ڑ��̃R�l�N�^

End Sub

Sub �O���b�h���ɑ�����()

    Dim sldidx As Integer       ' slide index
    Dim shprng As ShapeRange    ' shape range
    Dim shp As Shape            ' shape
    Dim shpcnt As Integer       ' shape count changed
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
    
    ' �P��X���C�h�`�F�b�N
    On Error GoTo ERROR_NO_ONE_SLIDE
    sldidx = ActiveWindow.Selection.SlideRange.SlideIndex
    On Error GoTo 0
    
    ' �}�`�I���`�F�b�N
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        If MsgBox( _
            "�I������Ă���}�`����܂���B" + vbCrLf + _
            "���ׂĂ̐}�`���O���b�h���ɑ����܂����H", _
            vbQuestion + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        Set shprng = ActivePresentation.Slides(sldidx).shapes.Range
    Else
        Set shprng = ActiveWindow.Selection.ShapeRange
    End If
    
    ' �I�����Ă���}�`�P�ʂňʒu��������
    ActiveWindow.Selection.Unselect
    ActivePresentation.Slides(sldidx).Select
    shpcnt = 0
    For Each shp In shprng
        
        ' �R�l�N�^����
        cnnct = False
        If shp.Connector Then
            If shp.ConnectorFormat.BeginConnected Then cnnct = True
            If shp.ConnectorFormat.EndConnected Then cnnct = True
        End If
        
        ' �R�l�N�^�ȊO�̏ꍇ
        If Not (cnnct) Then
               
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
        
    Next
    
    ' �������ʂ�ʒm
    If shpcnt = 0 Then
        MsgBox "���ׂĂ̐}�`���O���b�h���ɑ����Ă��܂��B", vbInformation
    End If
    
    Exit Sub
    
ERROR_NO_ONE_SLIDE:

    MsgBox "�X���C�h��I�����Ă�������"
    Exit Sub

End Sub
    
Sub �Б��ڑ��̃R�l�N�^()

    Dim sldidx As Integer       ' slide index
    Dim shprng As ShapeRange    ' shape range
    Dim shp As Shape            ' shape
    Dim flg As Boolean          ' flag
    
    ' �P��X���C�h�`�F�b�N
    On Error GoTo ERROR_NO_ONE_SLIDE
    sldidx = ActiveWindow.Selection.SlideRange.SlideIndex
    On Error GoTo 0
   
    ' �X���C�h�̐}�`�ꗗ
    ActiveWindow.Selection.Unselect
    ActivePresentation.Slides(sldidx).Select
    Set shprng = ActivePresentation.Slides(sldidx).shapes.Range
    For Each shp In shprng
        
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
    
    ' �Б��ڑ��̃R�l�N�^�������炩�������Ƃ�ʒm
    MsgBox "�Б��ڑ��̃R�l�N�^�͂���܂���B", vbInformation
    Exit Sub
    
ERROR_NO_ONE_SLIDE:

    MsgBox "�X���C�h��I�����Ă�������"
    Exit Sub

End Sub

