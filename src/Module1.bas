Attribute VB_Name = "Module1"
Option Explicit

Sub 図形をグリッドに揃えるCB(constrol As IRibbonControl)

    図形をグリッドに揃える

End Sub

Sub 図形をグリッドに揃える()

    Dim sldidx As Integer       ' slide index
    Dim shprng As ShapeRange    ' shape range
    Dim shp As Shape            ' shape
    Dim shpcnt As Integer       ' shape count changed
    Dim cnnct As Boolean    ' connector or not
    Dim left As Integer     ' left
    Dim top As Integer      ' top
    Dim width As Integer    ' width
    Dim height As Integer   ' height
    Dim lcnt As Integer     ' left count
    Dim lrem As Integer     ' left remain
    Dim tcnt As Integer     ' top count
    Dim trem As Integer     ' top remain
    Dim wcnt As Integer     ' width count
    Dim wrem As Integer     ' width remain
    Dim hcnt As Integer     ' height count
    Dim hrem As Integer     ' height remain
    
    ' 単一スライドチェック
    On Error GoTo ERROR_NO_ONE_SLIDE
    sldidx = ActiveWindow.Selection.SlideRange.SlideIndex
    On Error GoTo 0
    
    ' 図形選択チェック
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        If MsgBox( _
            "選択されている図形ありません。" + vbCrLf + _
            "すべての図形をグリッドに揃えますか？", _
            vbQuestion + vbOKCancel) = vbNo Then
            Exit Sub
        End If
        Set shprng = ActivePresentation.Slides(sldidx).shapes.Range
    Else
        Set shprng = ActiveWindow.Selection.ShapeRange
    End If
    
    ' 選択している図形単位で位置調整処理
    ActiveWindow.Selection.Unselect
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
            
            ' 中心に合わせて調整、グリッドの揃えた後で元に戻す
            left = left - ActivePresentation.PageSetup.SlideWidth / 2
            top = top - ActivePresentation.PageSetup.SlideHeight / 2
            
            ' 繰返数を求めた後で差分を計算しグリッドに揃うよう調整
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
                    CStr(shp.left) + " -> " + CStr(left) + vbCrLf + _
                    CStr(shp.top) + " -> " + CStr(top) + vbCrLf + _
                    CStr(shp.width) + " -> " + CStr(width) + vbCrLf + _
                    CStr(shp.height) + " -> " + CStr(height) + vbCrLf

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
    If shpcnt > 0 Then
        MsgBox _
            CStr(shpcnt) + "個の図形をグリッドに揃えました。" + vbCrLf + _
            "位置を調整した図形を選択しています。", _
            vbInformation
    Else
        MsgBox "位置を調整した図形はありません。", vbInformation
    End If
    
    Exit Sub
    
ERROR_NO_ONE_SLIDE:

    MsgBox "スライドを選択してください"
    Exit Sub

End Sub
    
