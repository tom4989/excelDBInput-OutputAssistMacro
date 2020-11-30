VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQL生成 
   Caption         =   "現在のシート"
   ClientHeight    =   7860
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12840
   OleObjectBlob   =   "frmSQL生成.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSQL生成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wb起動元ブック As Workbook
Public wb前回実行結果 As Workbook
Private txtSQL作成元シート名 As String
Private txtトランザクション識別文字列 As String
Private txtSQL実行バッチファイルパス As String
Private txtSQL実行ログファイルパス As String

Private Sub btnSQL生成_Click()

    'エラートラップ
    On Error GoTo ErrorCatch

    txbステータスバー.Value = get開始メッセージ("SQL生成")

    Dim objアプリケーション制御 As New clsアプリケーション制御
    objアプリケーション制御.アプリケーション制御切替 (False)

    wb起動元ブック.Activate
    
    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    
    Dim obj対象シート As Worksheet
    Set obj対象シート = ActiveSheet
    txtSQL作成元シート名 = ActiveSheet.Name
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    txtトランザクション識別文字列 = FSO.getBaseName(wb起動元ブック.Name) & "_" & txtSQL作成元シート名 & "_" & getTimestamp
    
    txtSQL.Value = ""
    frmSQL生成.Repaint
    
    ' SPOOL文の出力
    If ckbSPOOL Then
        txtSQL.Value = txtSQL.Value & vbCrLf & _
            "SPOOL """ & txtトランザクション識別文字列 & ".sql,log" & vbCrLf & vbCrLf
    
    If True Then
    
        txtSQL.Value = vbCrLf & "-- シート名：" & obj対象シート.Name & vbLf & _
            createSQL文(obj試験データシート, obj対象シート)
    Else
    
        Dim strSQL As String
        
        For Each obj対象シート In ActiveWorkbook.Sheets
        
            obj対象シート.Activate
            
            strSQL = strSQL & vbLf & "-- シート名：" & obj対象シート.Name
            strSQL = strSQL & createSQL文(obj試験データシート, obj対象シート)
            
        Next
        
        txtSQL.Value = strSQL
    
    End If
    
    txbステータスバー.Value = get終了メッセージ("SQL生成")
    
   ' 正常終了
   GoTo Finally
    
ErrorCatch:

    txbステータスバー.Value = get終了メッセージ("SQL生成")

Finally:

   '実行前の状態に戻す
   objアプリケーション制御.アプリケーション制御切替 (True)
        
End Sub

Private Function createSQL文( _
    ByRef obj試験データシート As cls試験データシート, _
    ByRef obj対象シート As Worksheet) As String
    
    If rdInsert Then
        createSQL文 = obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.INSERT文, rdb選択行のみ)
        
    ElseIf rdUpdate Then
        createSQL文 = obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.UPDATE文, rdb選択行のみ)
        
    ElseIf rdSELECT Then
        createSQL文 = obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.SELECT文, rdb選択行のみ)
        
    ElseIf rdDelete Then
        createSQL文 = obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.DELETE文, rdb選択行のみ)
        
    End If
    
End Function
        
Private Sub btnレコード取得_Click()
        
    Call レコード取得ボタン
        
    Unload frmSQL生成
        
End Sub
        
Private Sub レコード取得ボタン()

    txbステータスバー.Value = get開始メッセージ("レコード取得")
    
    'エラートラップ
    On Error GoTo ErrorCatch

    Dim objアプリケーション制御 As New clsアプリケーション制御
    objアプリケーション制御.アプリケーション制御切替 (False)
        
    ' 本処理
        
    wb起動元ブック.Activate
    
    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    
    Set wb前回実行結果 = obj試験データシート.getレコード(Nothing)
    
    If Not (wb前回実行結果 Is Nothing) Then
        Application.CutCopyMode = False
        
        wb前回実行結果.Activate
        wb前回実行結果.ActiveSheet.Range("A1").Select
        
        btnレコード追加取得.Enabled = True
        
    End If
        
    txbステータスバー.Value = get終了メッセージ("レコード取得")
        
   ' 正常終了
   GoTo Finally
    
ErrorCatch:

Finally:

    txbステータスバー.Value = get異常時メッセージ("レコード取得")

    '実行前の状態に戻す
    objアプリケーション制御.アプリケーション制御切替 (True)

End Sub

Private Sub btn結果をファイルに出力_Click()

    Dim obj設定シート As New cls設定シート

    mkdirIFNotExist obj設定シート.結果ファイル出力先

    Dim txt出力パス As String
    txt出力パス = obj設定シート.結果ファイル出力先 & "\" & txtSQL作成元シート名 & "_" & getTimestamp() & ".sql"

    Open txt出力パス For Output As #1

    Print #1, txtSQL.Value

    Close #1

    txbステータスバー = txt出力パス

End Sub

Private Sub btn更新前レコード取得_Click()

    レコード取得ボタン
    btn更新後レコード取得.Enabled = True

End Sub

Private Sub btn更新後レコード取得_Click()

    txbステータスバー.Value = get開始メッセージ("更新後レコード取得")
    
    'エラートラップ
    On Error GoTo ErrorCatch

    Dim objアプリケーション制御 As New clsアプリケーション制御
    objアプリケーション制御.アプリケーション制御切替 (False)
        
    ' 本処理
    wb起動元ブック.Activate
    
    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    
    Set wb前回実行結果 = obj試験データシート.getレコード(wb前回実行結果)
    
    If Not (wb前回実行結果 Is Nothing) Then
        Application.CutCopyMode = False
        
        wb前回実行結果.Activate
        wb前回実行結果.ActiveSheet.Range("A1").Select
        
        Call obj試験データシート.edit実行結果差分(wb前回実行結果)
    
    End If

    txbステータスバー.Value = get終了メッセージ("更新後レコード取得")

    btn更新後レコード取得.Enabled = False

   ' 正常終了
   GoTo Finally
    
ErrorCatch:

Finally:

    txbステータスバー.Value = get異常時メッセージ("レコード取得")

    '実行前の状態に戻す
    objアプリケーション制御.アプリケーション制御切替 (True)

End Sub

Private Sub UserForm_Initialize()

    Set wb起動元ブック = ActiveWorkbook
    
    btnSQL生成.SetFocus
    
End Sub

