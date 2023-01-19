Attribute VB_Name = "Module1"
Option Explicit

Sub グリッド線に揃える_onAction(constrol As IRibbonControl)

    グリッド線に揃える

End Sub

Sub 片側接続のコネクタ_onAction(constrol As IRibbonControl)

    片側接続のコネクタ

End Sub

Sub リンク切れのURL_onAction(constrol As IRibbonControl)

    リンク切れのURL

End Sub

Sub グリッド線に揃える()

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
    
    ' スライドがない場合をチェック
    Set test = Nothing
    On Error Resume Next
    Set test = ActiveWindow.Selection
    On Error GoTo 0
    If test Is Nothing Then
        Debug.Print "----" + vbCrLf + _
            "exit: cannot get selection"
        Exit Sub
    End If
    
    ' 図形未選択時はスライドマスタ以外の図形をすべて選択
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
    
        ' スライド取得
        Set sld = Nothing
        On Error Resume Next
        Set sld = ActivePresentation.Slides.FindBySlideID(ActivePresentation.Windows(1).Selection.SlideRange.SlideID)
        On Error GoTo 0
        If sld Is Nothing Then
            ' スライドの狭間を選択時はここで抜ける
            ' スライドマスタ選択時もここで抜ける
            ' ※本当はスライドマスタを取得できれば抜ける必要はない
            Debug.Print "----" + vbCrLf + _
                "exit: cannot get slide object"
            Exit Sub
        End If
        
        ' プレースホルダの図形IDの辞書化
        Set shpdic = CreateObject("Scripting.Dictionary")
        For Each shp In sld.Shapes.Placeholders
            shpdic(shp.Id) = shp.Id
        Next
    
        ' プレースホルダおよびフッターの図形以外を選択
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
        
        ' 選択した図形が0なら後続の処理でエラーになるため終了
        If shpcnt = 0 Then
            Exit Sub
        End If
                
    End If
    
    ' 選択図形を対象に処理開始
    Set shprng = ActiveWindow.Selection.ShapeRange
    ActiveWindow.Selection.Unselect
    
    ' 選択している図形単位で位置調整処理
    shpcnt = 0
    For Each shp In shprng
                       
        ' コネクタ判定
        cnnct = False
        If shp.Connector Then
            If shp.ConnectorFormat.BeginConnected Then cnnct = True
            If shp.ConnectorFormat.EndConnected Then cnnct = True
        End If
        
        ' コネクタ以外の場合
        If Not cnnct Then
            
            Dim is_locked As Boolean
            is_locked = False
            On Error Resume Next
            is_locked = shp.Locked
            On Error GoTo 0
            
            If is_locked Then
            
                ' ロックされている場合
                
                ' デバッグ：ロックの表示
                Debug.Print "----" + vbCrLf + _
                    "id     : " + CStr(shp.Id) + vbCrLf + _
                    "status : locked"
                        
            Else
                            
                ' ロックされていない場合
                            
                ' 作業用の位置を取得
                left = shp.left
                top = shp.top
                width = shp.width
                height = shp.height
                
                ' 中心に合わせて調整、グリッド線の揃えた後で元に戻す
                left = left - ActivePresentation.PageSetup.SlideWidth / 2
                top = top - ActivePresentation.PageSetup.SlideHeight / 2
                
                ' 繰返数を求めた後で差分を計算しグリッド線に揃うよう調整
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
                
                ' 元の位置に戻す
                left = left + ActivePresentation.PageSetup.SlideWidth / 2
                top = top + ActivePresentation.PageSetup.SlideHeight / 2
                
                If Abs(shp.left - left) < 0.01 And _
                    Abs(shp.top - top) < 0.01 And _
                    Abs(shp.width - width) < 0.01 And _
                    Abs(shp.height - height) < 0.01 Then
                    
                    ' 変更されていなければ何もしない
                Else
                
                    ' 変更されていれば新たに選択する
                    
                    ' デバッグ：変更内容の表示
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
    
Sub 片側接続のコネクタ()

    Dim sld As Slide    ' slide
    Dim shp As Shape    ' shape
    Dim flg As Boolean  ' flag
    
    ' スライド取得
    Set sld = Nothing
    On Error Resume Next
    Set sld = ActivePresentation.Slides.FindBySlideID(ActivePresentation.Windows(1).Selection.SlideRange.SlideID)
    On Error GoTo 0
    If sld Is Nothing Then
        ' スライドの狭間を選択時はここで抜ける
        ' スライドマスタ選択時もここで抜ける
        ' ※本当はスライドマスタを取得できれば抜ける必要はない
        Debug.Print "----" + vbCrLf + _
            "exit: cannot get slide object"
        Exit Sub
    End If
    
    ' スライドの図形一覧
    ActiveWindow.Selection.Unselect
    For Each shp In sld.Shapes.Range
        
        If shp.Connector Then
        
            flg = False
            
            ' 片側コネクタのチェック
            If shp.ConnectorFormat.BeginConnected And Not shp.ConnectorFormat.EndConnected Then flg = True
            If shp.ConnectorFormat.EndConnected And Not shp.ConnectorFormat.BeginConnected Then flg = True
            
            ' 片側接続のコネクタが見つかったら終了
            If flg Then
                shp.Select
                Exit Sub
            End If
            
        End If
        
    Next
    
    ' 片側接続のコネクタが見つからなかったことを通知
    MsgBox "片側接続のコネクタはありません。", vbInformation
    Exit Sub
    
End Sub

Sub リンク切れのURL()

    Dim strUrlOld As String
    Dim strUrlNew As String
    Dim lCnt As Integer
    Dim sld As Slide
    Dim shp As Shape
    Dim strMessage As String
    
    On Error GoTo ON_ERROR
    
    ' 現状の選択状態を解除
    ActiveWindow.Selection.Unselect
    
    ' スライドを先頭から列挙
    strUrlOld = ""
    strUrlNew = ""
    lCnt = 0
    For Each sld In ActivePresentation.Slides
    
        ' 処理中のスライドへ移動、.Selectでは表示は変わらない
        ActiveWindow.View.GotoSlide sld.SlideIndex
        
        ' スライド配下の図形を列挙
        For Each shp In sld.Shapes
        
            ' 図形をチェック
            ' - チェック失敗時はstrUrlOld等に値が設定
            ' - チェック失敗後は最後のスライドまでURLの再設定を実施
            CheckShape shp, strUrlOld, strUrlNew, lCnt
        Next
        
    Next
  
    ' スライド列挙後
    If lCnt = 0 Then
        ' チェック成功時
        ActiveWindow.Selection.Unselect
        strMessage = "最後まで検索しました。"
        Debug.Print "----" + vbCrLf + strMessage
        MsgBox strMessage, vbInformation
            
    Else
        ' チェック失敗時
        strMessage = "リンク切れのURLが見つかったため" + CStr(lCnt) + "件再設定しました。" + vbCrLf + _
            "再度実行してください。"
        Debug.Print "----" + vbCrLf + strMessage
        MsgBox strMessage, vbInformation

    End If
    
    Exit Sub
    
ON_ERROR:

    ' エラー時
    strMessage = Err.Source
    Debug.Print "----" + vbCrLf + strMessage
    MsgBox strMessage, vbExclamation

End Sub

Sub CheckShape(shp As Shape, _
    ByRef strUrlOld As String, _
    ByRef strUrlNew As String, _
    ByRef lCnt As Integer)

    ' 個々の図形をチェック
    
    ' stackoverflowを参考に実装
    ' https://stackoverflow.com/questions/55724877/how-to-obtain-shapes-to-hyperlinks-in-powerpoint-vba
    
    Dim mouseActivation As Variant
    Dim actionSetting As actionSetting
    Dim rangeLabel As TextRange
    Dim strUrl As String
    Dim strLabel As String
    Dim strType As String
    Dim i As Integer
    Dim j As Integer
    
    ' 図形を選択
    shp.Select
    
    If shp.Type = msoGroup Then
    
        ' 図形がグループの場合
        
        Dim shp2 As Shape
        For Each shp2 In shp.GroupItems
            CheckShape shp2, strUrlOld, strUrlNew, lCnt
        Next
    
    Else
    
        ' 図形がグループでない場合
        
        For Each actionSetting In shp.ActionSettings
        
            ' *** 図形のリンク先をチェック ***
            If actionSetting.Action = ppActionHyperlink Then
                Set rangeLabel = Nothing
                strUrl = actionSetting.Hyperlink.Address
                strLabel = ""
                strType = "図形"
                CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
            End If
            
        Next
    
        ' *** リンクされた図形のリンク先をチェック ***
        If shp.Type = msoLinkedPicture Then
            Set rangeLabel = Nothing
            strUrl = shp.LinkFormat.SourceFullName
            strLabel = ""
            strType = "リンクされた図形"
            CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
        ElseIf shp.TextFrame.HasText Then
                    
            For Each mouseActivation In Array(ppMouseClick, ppMouseOver)
                Set actionSetting = shp.TextFrame.TextRange.ActionSettings(mouseActivation)
                
                ' *** テキスト全体のリンク先をチェック ***
                If actionSetting.Action = ppActionHyperlink Then
                    Set rangeLabel = shp.TextFrame.TextRange.Characters(1, shp.TextFrame.TextRange.Characters.Count)
                    strUrl = actionSetting.Hyperlink.Address
                    strLabel = shp.TextFrame.TextRange.Text
                    strType = "テキスト"
                    CheckShapeSub shp, strUrlOld, strUrlNew, lCnt, actionSetting, rangeLabel, strUrl, strLabel, strType
                Else
                                        
                    strUrl = ""
                    For i = 1 To shp.TextFrame.TextRange.Characters.Count
                        Set actionSetting = shp.TextFrame.TextRange.Characters(i).ActionSettings(mouseActivation)
                        
                        ' *** テキストのリンク先をチェック ***
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
                            strType = "テキスト*"
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
                strUrlNew = InputBox(strType + "の新しいURLを入力してください。" + vbCrLf + strUrl, strStatus, strUrl)
                If strUrlNew = "" Then
                    If MsgBox("このURLをスキップして続行しますか？" + vbCrLf + "はい：スキップして続行" + vbCrLf + "いいえ：終了", vbYesNo + vbQuestion) = vbYes Then
                        Exit Sub
                    Else
                        Err.Raise 1001, "キャンセルされました。"
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
    
        ' 以下のサイトを参考に実装
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


