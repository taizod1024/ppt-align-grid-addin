Attribute VB_Name = "Module1"
Option Explicit

Sub グリッド線に揃える_onAction(constrol As IRibbonControl)

    グリッド線に揃える

End Sub

Sub 片側接続のコネクタ_onAction(constrol As IRibbonControl)

    片側接続のコネクタ

End Sub

Sub グリッド線に揃える()

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
    
    ' 単一スライドチェック
    On Error GoTo ERROR_NO_ONE_SLIDE
    sldidx = ActiveWindow.Selection.SlideRange.SlideIndex
    On Error GoTo 0
    
    ' 図形選択チェック
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        If MsgBox( _
            "選択されている図形ありません。" + vbCrLf + _
            "すべての図形をグリッド線に揃えますか？", _
            vbQuestion + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        Set shprng = ActivePresentation.Slides(sldidx).shapes.Range
    Else
        Set shprng = ActiveWindow.Selection.ShapeRange
    End If
    
    ' 選択している図形単位で位置調整処理
    ActiveWindow.Selection.Unselect
    ActivePresentation.Slides(sldidx).Select
    shpcnt = 0
    For Each shp In shprng
        
        ' コネクタ判定
        cnnct = False
        If shp.Connector Then
            If shp.ConnectorFormat.BeginConnected Then cnnct = True
            If shp.ConnectorFormat.EndConnected Then cnnct = True
        End If
        
        ' コネクタ以外の場合
        If Not (cnnct) Then
               
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
    
    ' 調整結果を通知
    If shpcnt = 0 Then
        MsgBox "すべての図形がグリッド線に揃っています。", vbInformation
    End If
    
    Exit Sub
    
ERROR_NO_ONE_SLIDE:

    MsgBox "スライドを選択してください"
    Exit Sub

End Sub
    
Sub 片側接続のコネクタ()

    Dim sldidx As Integer       ' slide index
    Dim shprng As ShapeRange    ' shape range
    Dim shp As Shape            ' shape
    Dim flg As Boolean          ' flag
    
    ' 単一スライドチェック
    On Error GoTo ERROR_NO_ONE_SLIDE
    sldidx = ActiveWindow.Selection.SlideRange.SlideIndex
    On Error GoTo 0
   
    ' スライドの図形一覧
    ActiveWindow.Selection.Unselect
    ActivePresentation.Slides(sldidx).Select
    Set shprng = ActivePresentation.Slides(sldidx).shapes.Range
    For Each shp In shprng
        
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
    
    ' 片側接続のコネクタが見つからかったことを通知
    MsgBox "片側接続のコネクタはありません。", vbInformation
    Exit Sub
    
ERROR_NO_ONE_SLIDE:

    MsgBox "スライドを選択してください"
    Exit Sub

End Sub

