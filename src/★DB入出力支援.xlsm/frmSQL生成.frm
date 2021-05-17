VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQL生成 
   Caption         =   "現在のシート"
   ClientHeight    =   8388
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
Private txtSQL作成日時 As String
Private txt結果ファイル出力先 As String

Private obj設定値シート As cls設定値シート

' *********************************************************************************************************************
' * 機能　：フォーム生成時処理
' *********************************************************************************************************************
'
Private Sub UserForm_Initialize()

    Set wb起動元ブック = ActiveWorkbook
    
    btnSQL生成.SetFocus

End Sub

' *********************************************************************************************************************
' * 機能　：フォーム生成時処理
' *********************************************************************************************************************
'
Public Sub 設定値ロード(arg設定値シート As cls設定値シート)

    Set obj設定値シート = arg設定値シート
    
    Dim txt接続情報 As Variant
    
    If Not obj設定値シート Is Nothing Then
        For Each txt接続情報 In obj設定値シート.設定値リスト.Item("接続情報")
    
            cmb接続情報.AddItem (txt接続情報)
        
        Next txt接続情報
        
        cmb接続情報.ListIndex = 0
    
        txt結果ファイル出力先 = Replace(obj設定値シート.設定値リスト.Item("結果ファイル出力先"), _
            "%USERPROFILE%", Environ("UserProfile"))

    End If
    
End Sub

' *********************************************************************************************************************
' * 機能　：レコード取得ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btnレコード取得_Click()
        
    Call レコード取得ボタン
        
    Unload frmSQL生成
        
End Sub

' *********************************************************************************************************************
' * 機能　：レコード取得処理
' *********************************************************************************************************************
'
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
    Call obj試験データシート.初期化(obj設定値シート, cmb接続情報.Text)
    
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

' *********************************************************************************************************************
' * 機能　：更新前レコード取得ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btn更新前レコード取得_Click()

    レコード取得ボタン
    btn更新後レコード取得.Enabled = True

End Sub

' *********************************************************************************************************************
' * 機能　：更新後レコード取得ボタン押下時の処理
' *********************************************************************************************************************
'
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

' *********************************************************************************************************************
' * 機能　：SQL生成ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btnSQL生成_Click()

    'エラートラップ
    On Error GoTo ErrorCatch

    txtSQL作成日時 = getTimestamp()

    txbステータスバー.Value = get開始メッセージ("SQL生成")

    Dim objアプリケーション制御 As New clsアプリケーション制御
    objアプリケーション制御.アプリケーション制御切替 (False)

    wb起動元ブック.Activate
    
    Dim obj試験データシート As cls試験データシート
    Set obj試験データシート = New cls試験データシート
    Call obj試験データシート.初期化(obj設定値シート, cmb接続情報.Text)
    
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
    End If
    
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
    
    ' 後続のボタンをenableに変更
    btnクリップボードにコピー.Enabled = True
    btn結果をファイルに出力.Enabled = True
    btn出力ファイルを実行.Enabled = False
    btnエラー確認.Enabled = False


    txbステータスバー.Value = get終了メッセージ("SQL生成")
    
   ' 正常終了
   GoTo Finally
    
ErrorCatch:

    txbステータスバー.Value = get終了メッセージ("SQL生成")

Finally:

   '実行前の状態に戻す
   objアプリケーション制御.アプリケーション制御切替 (True)
        
End Sub

' *********************************************************************************************************************
' * 機能　：SQL文生成処理
' *********************************************************************************************************************
'
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
        
    ElseIf rdDELETEINSERT Then
        
        createSQL文 = obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.DELETE文, rdb選択行のみ) & _
            obj試験データシート.対象シートSQL文作成(obj対象シート, SQL種別.INSERT文, rdb選択行のみ)
        
    End If
    
End Function

' *********************************************************************************************************************
' * 機能　：クリップボードにコピー
' *********************************************************************************************************************
'
Private Sub btnクリップボードにコピー_Click()

    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = txtSQL.Value
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
    
    txbステータスバー.Value = get終了メッセージ("クリップボードにコピー")

End Sub

' *********************************************************************************************************************
' * 機能　：結果をファイルに出力ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btn結果をファイルに出力_Click()

    mkdirIFNotExist txt結果ファイル出力先

    Dim txt出力パス As String
    txt出力パス = txt結果ファイル出力先 & "\" & txtSQL作成元シート名 & "_" & txtSQL作成日時 & ".sql"

    Open txt出力パス For Output As #1
    
    Dim txtSQLファイル内容 As String
    txtSQLファイル内容 = "SPOOL " & txtSQL作成元シート名 & "_" & txtSQL作成日時 & ".log" & vbCrLf & _
        txtSQL.Value & vbCrLf & _
        "SPOOL OFF" & vbCrLf & _
        "EXIT"
    
    Print #1, txtSQLファイル内容
    Close #1

    Dim txtバッチファイル内容 As String
    
    txtバッチファイル内容 = "cd " & txt結果ファイル出力先 & vbCrLf
    
    If obj設定値シート.設定値リスト.Exists("ORACLE_HOME") Then
     
        txtバッチファイル内容 = txtバッチファイル内容 & _
            "set ORACLE_HOME=" & obj設定値シート.設定値リスト("ORACLE_HOME") & vbCrLf
    End If
    
    txtバッチファイル内容 = txtバッチファイル内容 & _
        "sqlplus " & _
        obj設定値シート.設定値リスト("接続情報").Item(cmb接続情報.Value).Item("UID") & "/" & _
        obj設定値シート.設定値リスト("接続情報").Item(cmb接続情報.Value).Item("PWD") & "@" & _
        obj設定値シート.設定値リスト("接続情報").Item(cmb接続情報.Value).Item("DSN") & _
        " @" & txt出力パス & vbCrLf & "pause"

    Open txt出力パス & ".bat" For Output As #1
    Print #1, txtバッチファイル内容
    Close #1

    ' 結果出力
    txbステータスバー.Value = get終了メッセージ("結果をファイルに出力")
    txbステータスバー = txbステータスバー.Value & vbCr & txt出力パス

    btn出力ファイルを実行.Enabled = True

End Sub

' *********************************************************************************************************************
' * 機能　：結果をファイルに出力ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btn出力ファイルを実行_Click()

    Dim txt出力パス As String
    txt出力パス = txt結果ファイル出力先 & "\" & txtSQL作成元シート名 & "_" & txtSQL作成日時 & ".sql.bat"

    Call Shell(txt出力パス, vbNormalFocus)

    btnエラー確認.Enabled = True
    
End Sub

' *********************************************************************************************************************
' * 機能　：エラー確認ボタン押下時の処理
' *********************************************************************************************************************
'
Private Sub btnエラー確認_Click()

    Dim txt出力パス As String
    txt出力パス = txt結果ファイル出力先 & "\" & txtSQL作成元シート名 & "_" & txtSQL作成日時 & ".log"

    Call Shell("notepad.exe " & txt出力パス, vbNormalFocus)

End Sub
