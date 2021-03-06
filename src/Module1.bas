Attribute VB_Name = "Module1"
Option Explicit

Sub グリッド線に揃える_onAction(constrol As IRibbonControl)

    グリッド線に揃える

End Sub

Sub 片側接続のコネクタ_onAction(constrol As IRibbonControl)

    片側接続のコネクタ

End Sub

Sub グリッド線に揃える()

    Dim test As Object          ' test object
    Dim sld As slide            ' slide
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
        For Each shp In sld.shapes.Placeholders
            shpdic(shp.Id) = shp.Id
        Next
    
        ' プレースホルダおよびフッターの図形以外を選択
        shpcnt = 0
        For Each shp In sld.shapes.Range
        
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
               
            ' 作業用の位置を取得
            left = shp.left
            top = shp.top
            width = shp.width
            height = shp.height
            
            ' 中心に合わせて調整、グリッド線の揃えた後で元に戻す
            left = left - ActivePresentation.PageSetup.SlideWidth / 2
            top = top - ActivePresentation.PageSetup.SlideHeight / 2
            
            ' 繰返数を求めた後で差分を計算しグリッド線に揃うよう調整
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
    
    Exit Sub
    
End Sub
    
Sub 片側接続のコネクタ()

    Dim sld As slide    ' slide
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
    For Each shp In sld.shapes.Range
        
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

