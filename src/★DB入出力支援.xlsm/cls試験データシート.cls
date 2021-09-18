VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls試験データシート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

Const cnst上限レコード取得数 = 10000 ' 環境に合わせて増減

Const cnst試験データ開始行 = 5
Const cnstヘッダ行数 = 5
Const cnstテーブル物理名記載列 = 4

Const cnstテーブル名記載行idx = 1
Const cnstデータ開始行idx = 2
Const cnstデータ終了行idx = 3
Const cnstデータ終了列idx = 4

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
Public Enum SQL種別
    SELECT文 = 1
    INSERT文 = 2
    UPDATE文 = 3
    DELETE文 = 4
    DELETEINSERT文 = 5
End Enum

Private oraconn As Object

Private dic設定値 As Object

Private txtDBMS As String

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

    Set oraconn = CreateObject("ADODB.Connection")
        
End Sub


' *********************************************************************************************************************
' 機能：初期化
' *********************************************************************************************************************
'
Public Sub 初期化(arg設定値シート As cls設定値シート, argDBMS As String)

    Set dic設定値 = arg設定値シート.設定値リスト

    ' DBMSの種類を設定
    txtDBMS = argDBMS

    Dim dic接続情報 As Object
    Set dic接続情報 = dic設定値.Item("接続情報").Item(txtDBMS)
    
    Dim txt接続文字列
    
    Dim var設定値 As Variant
    
    For Each var設定値 In dic接続情報
    
        txt接続文字列 = txt接続文字列 & var設定値 & "=" & dic接続情報.Item(var設定値) & ";"
        
    Next var設定値
    
    ' データベースに接続する
    oraconn.ConnectionString = txt接続文字列

End Sub

' *********************************************************************************************************************
' 機能：終了時の処理
' *********************************************************************************************************************
'
Public Sub Class_Terminate()
    Set oraconn = Nothing
End Sub

' *********************************************************************************************************************
' 機能：シート上、オープン状態のテーブルのレコードをDBに取得しに行く
' *********************************************************************************************************************
'
Public Function getレコード(ByRef wb前回実行結果 As Workbook) As Workbook

    oraconn.Open
    
    If executeCountSQL Then ' ユーザが件数を確認し続行した場合のみデータを取得
    
        Set getレコード = executeSQL(wb前回実行結果)

    End If
    
    oraconn.Close
    Set oraconn = Nothing
    
    Exit Function
        
End Function
    
' *********************************************************************************************************************
' 機能：COUNTのSQL発行を行う。
'       カウントの結果を参照したユーザが、処理を中断した場合、FALSEを返す
' *********************************************************************************************************************
'
Private Function executeCountSQL() As Boolean

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    Dim obj試験データテーブル() As cls試験データテーブル
    obj試験データテーブル = getテーブル座標情報
    
    With ActiveSheet
    
        Dim i As Long
        
        For i = LBound(obj試験データテーブル) To UBound(obj試験データテーブル)
        
            ' ---------------------------------------------------------------------------------------------------------
            ' SQL文の作成
            ' ---------------------------------------------------------------------------------------------------------
            
            ' ヘッダ部分が非表示の場合はスキップする
            If obj試験データテーブル(i).Hidden() Then
            
                GoTo continue
            End If
        
            Dim strQuery As String
            strQuery = obj試験データテーブル(i).createSELECT文From複数行()
            
            ' ---------------------------------------------------------------------------------------------------------
            ' SQL文の発行
            ' ---------------------------------------------------------------------------------------------------------
            
            Dim txtCountQuery As String
            txtCountQuery = obj試験データテーブル(i).createCount文(strQuery)
            
            log ("発行するSQL：" & txtCountQuery)
            Set rs = oraconn.Execute(txtCountQuery)
            
            obj試験データテーブル(i).setDBCount結果 (rs.Fields("COUNT"))
    
continue:
    
        Next i
        
    End With
    
    ' レコードセットをクローズする
    rs.Close
    Set rs = Nothing
    
    ' COUNTモードで実行した場合、取得予定の件数を利用者に通知する。
    ' 尚、件数が上限を超えるような場合は、利用者にその旨を通知し、処理を打ち切る。
    
    Dim txtCountList As String
    Dim lngレコード総数 As Long: lngレコード総数 = 0
    
    For i = LBound(obj試験データテーブル) To UBound(obj試験データテーブル)
    
        If Not obj試験データテーブル(i).Hidden Then
        
            txtCountList = txtCountList & _
                obj試験データテーブル(i).getテーブル論理名 & _
                "：" _
                & obj試験データテーブル(i).getDBCount結果 & " 件" & vbCrLf
                
            lngレコード総数 = lngレコード総数 + obj試験データテーブル(i).getDBCount結果
        End If
        
    Next i
    
    Dim txtMsg As String
    
    If lngレコード総数 <= cnst上限レコード取得数 Then
    
        txtMsg = "総件数 " & lngレコード総数 & " 件のレコードを結果として取得します。" _
            & vbCrLf & vbCrLf & txtCountList
            
        If MsgBox(txtMsg, vbOKCancel) <> vbOK Then
        
            executeCountSQL = False
            Exit Function
        End If
    
    Else
    
        txtMsg = "総件数（" & lngレコード総数 _
            & " 件）が上限件数（" & cnst上限レコード取得数 & " 件）を超えています。" _
            & "条件を見直して下さい。" _
            & vbCrLf & vbCrLf & txtCountList
            
        MsgBox (txtMsg)
        
        executeCountSQL = False
        Exit Function
        
     End If
        
    executeCountSQL = True
        
End Function


' *********************************************************************************************************************
' 機能：COUNTのSQL発行を行う。
'       カウントの結果を参照したユーザが、処理を中断した場合、FALSEを返す
' *********************************************************************************************************************
'
Private Function executeSQL(ByRef wb前回実行結果 As Workbook) As Workbook

    Dim wb実行元ブック As Workbook
    Set wb実行元ブック = ActiveWorkbook
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    Dim obj試験データテーブル() As cls試験データテーブル
    obj試験データテーブル = getテーブル座標情報
    
    Dim lng出力最終行 As Long
    lng出力最終行 = 5
    
    ' ブックの準備
    If wb前回実行結果 Is Nothing Then
    
        Dim wb実行結果 As Workbook
        Set wb実行結果 = Workbooks.Add
        
        wb実行元ブック.ActiveSheet.Copy Before:=wb実行結果.Sheets(1)
        
    Else
        Set wb実行結果 = wb前回実行結果
        
        wb実行元ブック.ActiveSheet.Copy After:=wb実行結果.Sheets(wb実行結果.Sheets.Count)
    
    End If
    
    Dim txt実行結果シート名 As String
    txt実行結果シート名 = Format(Now(), "yyyymmdd_HHnnss")
    
    wb実行結果.Sheets(wb実行元ブック.ActiveSheet.Name).Name = txt実行結果シート名
    
    If wb前回実行結果 Is Nothing Then
        Call 不要シート削除(wb実行結果, txt実行結果シート名)
    End If
    
    wb実行結果.ActiveSheet.Range("A" & lng出力最終行, ActiveCell.SpecialCells(xlLastCell)).Delete
    wb実行結果.ActiveSheet.Outline.ShowLevels RowLevels:=2

    wb実行結果.ActiveSheet.Rows.Ungroup
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1

    wb実行元ブック.Activate
    
    With ActiveSheet
    
        Dim i As Long
        
        For i = LBound(obj試験データテーブル) To UBound(obj試験データテーブル)
        
            ' ---------------------------------------------------------------------------------------------------------
            ' SQL文の作成
            ' ---------------------------------------------------------------------------------------------------------
            
            ' ヘッダ部分が非表示の場合はスキップする
            If obj試験データテーブル(i).Hidden() Then
            
                GoTo continue
            End If
            
            
            Dim txtQuery As String
            txtQuery = obj試験データテーブル(i).createSELECT文From複数行()
            
            ' ---------------------------------------------------------------------------------------------------------
            ' SQL文の発行
            ' ---------------------------------------------------------------------------------------------------------
                        
            Dim txtCountQuery As String
            txtCountQuery = obj試験データテーブル(i).createCount文(txtQuery)
            
            log ("発行するSQL：" & txtCountQuery)
            Set rs = oraconn.Execute(txtCountQuery)
            
            obj試験データテーブル(i).setDBCount結果 (rs.Fields("COUNT"))
            
            rs.Close
            Set rs = Nothing
            
            log ("発行するSQL：" & txtQuery)
            Set rs = oraconn.Execute(txtQuery)
            
            obj試験データテーブル(i).copy枠Toクリップボード
            
            With wb実行結果.ActiveSheet.Range("A" & lng出力最終行)
            
                .PasteSpecial ' データ行をコピー
                
            End With
            
            With wb実行結果.ActiveSheet.Range("G" & lng出力最終行)
            
                .ClearComments
                .AddComment ("-- 結果取得時のSQL" & vbCrLf _
                    & SQL整形(txtQuery)) ' テーブル論理名のセルにデータ取得に使用したSQLをコメントで残す
                .Comment.Shape.TextFrame.AutoSize = True
                
                Dim lngデータ行数 As Long
                
                If obj試験データテーブル(i).getDBCount結果 = 0 Then
                
                    lngデータ行数 = 1
                Else
                    lngデータ行数 = obj試験データテーブル(i).getDBCount結果
                End If
                
                ' 件数の情報を書き込み
                .Value = "=COUNTA(B" & lng出力最終行 + cnstヘッダ行数 & ":B" & lng出力最終行 + cnstヘッダ行数 + lngデータ行数 & ")"
                
            End With
            
            lng出力最終行 = lng出力最終行 + cnstヘッダ行数
            
            ' 行のクリア（記入された文字＋セルの背景色）
            wb実行結果.ActiveSheet.Rows(lng出力最終行).ClearContents
            
            With wb実行結果.ActiveSheet.Rows(lng出力最終行).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            If obj試験データテーブル(i).getDBCount結果 > 0 Then
            
                ' 行の罫線をコピー
                wb実行結果.ActiveSheet.Rows(lng出力最終行).Copy
                
                wb実行結果.ActiveSheet.Rows(lng出力最終行 & ":" _
                    & lng出力最終行 + obj試験データテーブル(i).getDBCount結果() - 1).PasteSpecial
                    
                    Dim var取得結果 As Variant
                    ReDim var取得結果( _
                        1 To obj試験データテーブル(i).getDBCount結果, _
                        1 To rs.Fields.Count)
                        
                    Dim j As Long
                    
                    For j = 1 To UBound(var取得結果)
                    
                        If rs.EOF Then
                        
                            Exit For
                        End If
                        
                        Dim k As Long
                        
                        For k = 1 To rs.Fields.Count
                        
                            var取得結果(j, k) = rs.Fields(k - 1)
                            
                        Next k
                         
                        rs.MoveNext
                        
                    Next j
                    
                    With wb実行結果.ActiveSheet
                    
                        .Range(.Cells(lng出力最終行, 2), _
                            .Cells(lng出力最終行 + UBound(var取得結果) - 1, UBound(var取得結果, 2) + 1)) = var取得結果
                            
                    End With
            
                    lng出力最終行 = lng出力最終行 + obj試験データテーブル(i).getDBCount結果() + 1
                    
                Else
                
                    lng出力最終行 = lng出力最終行 + 2 ' 空枠＋余白で2行
                    
                End If
                
continue:

            Next i
            
        End With
        
        ' レコードセットをクローズする
        rs.Close
        Set rs = Nothing
        
        Set executeSQL = wb実行結果
            
            
End Function

' *********************************************************************************************************************
' 機能：対象シートのSQL文生成
' *********************************************************************************************************************
'
Public Function 対象シートSQL文作成( _
    ByRef obj対象シート As Worksheet, ByVal argSQL種別 As SQL種別, Optional ByVal is選択行のみ As Boolean = False) As String

    ' If obj対象シート.Range("A1") <> "凡例" Then
    '     Exit Function
    ' End If
    
    Dim obj試験データテーブル() As cls試験データテーブル
    obj試験データテーブル = getテーブル座標情報
    
    Dim stb結果 As New clsStringBuilder
    
    Dim i As Long
    
    For i = LBound(obj試験データテーブル) To UBound(obj試験データテーブル)
    
        If argSQL種別 = SQL種別.SELECT文 Then
            stb結果.append (obj試験データテーブル(i).createSELECT文From複数行To複数SQL(is選択行のみ))
            
        ElseIf argSQL種別 = SQL種別.INSERT文 Then
            stb結果.append (obj試験データテーブル(i).createInsert文(obj対象シート, is選択行のみ))
        
        ElseIf argSQL種別 = SQL種別.UPDATE文 Then
            stb結果.append (obj試験データテーブル(i).createUpdate文(obj対象シート, is選択行のみ))
        
        ElseIf argSQL種別 = SQL種別.DELETE文 Then
           stb結果.append (obj試験データテーブル(i).createDelete文(obj対象シート, is選択行のみ))
        
        End If
        
    Next i
    
    対象シートSQL文作成 = stb結果.toString
    
End Function

' *********************************************************************************************************************
' 機能：テーブルのグループごとの表示/非表示を制御する
' *********************************************************************************************************************
'
Public Sub 表示非表示の切替(Optional ByVal txtテーブル名 As String = "*マスタ")

    Dim varテーブル座標情報 As Variant
    varテーブル座標情報 = getテーブル座標情報
    
    Dim i As Long
    
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    ActiveSheet.Rows.Ungroup
    
    For i = LBound(varテーブル座標情報) To UBound(varテーブル座標情報)
    
        If Not ActiveSheet.Range("A" & varテーブル座標情報(i, 1)) Like txtテーブル名 Then
        
            Range(Cells(varテーブル座標情報(i, 1) + 1, 1), Cells(varテーブル座標情報(i, 3), 1)).Rows.Group
            
        End If
    Next i
    
    ActivceSheet.Outline.ShowLevels RowLevels:=1
    
    For i = LBound(varテーブル座標情報) To UBound(varテーブル座標情報)
    
        If ActiveSheet.Range("A" & varテーブル座標情報(i, 1)) Like txtテーブル名 Then
        
        
            Range(Cells(varテーブル座標情報(i, 1) + 1, 1), Cells(varテーブル座標情報(i, 3), 1)).Row.Group
        
        End If
    Next i

End Sub


' *********************************************************************************************************************
' 機能：引数で指定されたブックの後方2シートを比較し、差分に関する情報を後方のシートに付与する。
' *********************************************************************************************************************
'
Public Sub edit実行結果差分(ByRef wb引数実行結果 As Workbook)

    wb引数実行結果.Activate
    
    
    Dim ws実行前, ws実行後 As Worksheet
    
    Set ws実行前 = wb引数実行結果.Sheets(wb引数実行結果.Sheets.Count - 1)
    Set ws実行後 = wb引数実行結果.Sheets(wb引数実行結果.Sheets.Count)
    
    Dim obj実行前データテーブル() As cls試験データテーブル
    Dim obj実行後データテーブル() As cls試験データテーブル

    ws実行前.Activate
    obj実行前データテーブル = getテーブル座標情報()
    
    ws実行後.Activate
    obj実行後データテーブル = getテーブル座標情報()
    
    Dim i, j As Long
    i = UBound(obj実行後データテーブル)
    j = UBound(obj実行前データテーブル)
    
    ' 実行後を軸に、実行前とのマッチングを行っていく
    Do While (UBound(obj実行後データテーブル) <= i Or LBound(obj実行前データテーブル) <= j)
    
        If obj実行後データテーブル(i).getテーブル物理名 <> obj実行前データテーブル(j).getテーブル物理名 Then
        
            j = j - 1
            
            If j < UBound(obj実行前データテーブル) Then
            
                Exit Sub ' 実行前テーブルが存在しなくなった場合、終了
            End If
            
        Else
            Call compareテーブル(obj実行前データテーブル(j), obj実行後データテーブル(i))
            i = i - 1
            j = j - 1
            
        End If
        
    Loop

End Sub

' =====================================================================================================================
' Privateメソッド
' =====================================================================================================================
'

' *********************************************************************************************************************
' 機能：テーブル内容の比較を行う
' *********************************************************************************************************************
'
Private Sub compareテーブル( _
    ByRef obj実行前データテーブル As cls試験データテーブル, ByRef obj実行後データテーブル As cls試験データテーブル)
    
    If obj実行前データテーブル.get件数() = 0 And obj実行後データテーブル.get件数() = 0 Then
    
        Exit Sub ' 両方0件の場合、何もしない
    End If
        
    Dim i, j As Long
    
    With obj実行後データテーブル
    
        i = .データ終了行
        j = obj実行前データテーブル.データ終了行
        
        Do While (i >= .データ開始行 Or j >= obj実行前データテーブル.データ開始行)
        
            ' データがない行を読み飛ばす
            If .get主キー項目連結文字列(i) = "" Then
            
                i = i - 1
                GoTo continue
                
            ElseIf obj実行前データテーブル.get主キー項目連結文字列(j) = "" Then
            
                j = j - 1
                GoTo continue
                
            End If
            
            ' いずれかのデータを読み切っている場合
            If j < obj実行前データテーブル.データ開始行 Then
                
                Call log差異("読切追加", j, obj実行前データテーブル, i, obj実行後データテーブル)
                obj実行後データテーブル.edit選択行強調 (i)
                
                With .対象シート.Cells(i, 1)
                    .AddComment ("■追加")
                    .Comment.Shape.TextFrame.AutoSize = True
                End With
                
                i = i - 1
                
                GoTo continue
                
            ElseIf i < .データ開始行 Then
            
                Call log差異("読切削除", j, obj実行前データテーブル, i, obj実行後データテーブル)
                obj実行後データテーブル.add空行 (obj実行後データテーブル.データ開始行)
                obj実行後データテーブル.edit選択行網掛 (obj実行後データテーブル.データ開始行)
            
                With .対象シート.Cells(obj実行後データテーブル.データ開始行, 1)
                
                    .AddComment ("■削除" & vbCrLf _
                        & Join(obj実行前データテーブル.get主キー(), " ") & vbLf _
                        & Join(obj実行前データテーブル.get主キー(j), " "))
                        
                    .Comment.Shape.TextFrame.AutoSize = True
                    
                End With
                
                j = j - 1
                
                GoTo continue
                
            End If
            
            ' まだ、いずれのデータも残っている場合
            If .get主キー項目連結文字列(i) = obj実行前データテーブル.get主キー項目連結文字列(j) Then
            
                Call log差異("比較一致", j, obj実行前データテーブル, i, obj実行後データテーブル)
            
                Call compareレコード(obj実行前データテーブル, obj実行後データテーブル, j, i)
                
                i = i - 1
                j = j - 1
                
            ElseIf StrComp(.get主キー項目連結文字列(i), obj実行前データテーブル.get主キー項目連結文字列(j), _
                vbBinaryCompare) > 0 Then
                
                Call log差異("比較追加", j, obj実行前データテーブル, i, obj実行後データテーブル)
                obj実行後データテーブル.edit選択行強調 (i)
                
                With .対象シート.Cells(i, 1)
                    .AddComment ("■追加")
                    .Comment.Shape.TextFrame.AutoSize = True
                End With
                
                i = i - 1
                
            Else
                Call log差異("比較削除", j, obj実行前データテーブル, i, obj実行後データテーブル)
                obj実行後データテーブル.add空行 (i + 1)
                obj実行後データテーブル.edit選択行網掛 (i + 1)
                
                With .対象シート.Cells(i + 1, 1)
                
                    .AddComment ("■削除" & vbCrLf _
                        & Join(obj実行前データテーブル.get主キー(), " ") & vbCrLf _
                        & Join(obj実行前データテーブル.get主キー(j), " "))
                        
                    .Comment.Shape.TextFrame.AutoSize = True
                
                End With
                
                j = j - 1
                
            End If
continue:
        Loop
        
    End With
    
End Sub


' *********************************************************************************************************************
' 機能：デバッグ用の差異情報出力メソッド
' *********************************************************************************************************************
'
Private Sub compareレコード( _
    ByRef obj実行前データテーブル As cls試験データテーブル, _
    ByRef obj実行後データテーブル As cls試験データテーブル, _
    ByVal lng実行前対象行 As Long, ByVal lng実行後対象行 As Long)
    
    Dim txtコメント文字列 As String
    
    Dim txt変更前, txt変更後 As String
    
    Dim i As Long
    For i = 2 To obj実行前データテーブル.カラム終了列
    
        txt変更前 = obj実行前データテーブル.対象シート.Cells(lng実行前対象行, i)
        txt変更後 = obj実行後データテーブル.対象シート.Cells(lng実行後対象行, i)
        
        
        If txt変更前 <> txt変更後 Then
        
            Call obj実行後データテーブル.edit変更強調色(obj実行後データテーブル.対象シート.Cells(lng実行後対象行, i))
            
            txtコメント文字列 = txtコメント文字列 & vbCrLf _
                & obj実行前データテーブル.getカラム論理名(i) & "：" & txt変更前 & " → " & txt変更後
        End If
    Next i
    
    If txtコメント文字列 <> "" Then
        With obj実行後データテーブル.対象シート.Cells(lng実行後対象行, 1)
        
            .AddComment ("■更新" & txtコメント文字列)
            .Comment.Shape.TextFrame.AutoSize = True
            
        End With
    End If
End Sub

' *********************************************************************************************************************
' 機能：デバッグ用の差異情報出力メソッド
' *********************************************************************************************************************
'
Private Sub log差異(ByVal txtメッセージ As String, _
    ByVal lng変更前行番号 As Long, ByRef obj実行前データテーブル As cls試験データテーブル, _
    ByVal lng変更後行番号 As Long, ByRef obj実行後データテーブル As cls試験データテーブル)
    
    log obj実行前データテーブル.getテーブル論理名 & vbTab & txtメッセージ _
        & vbTab _
        & "変更前" & vbTab & lng変更前行番号 & "行目" & vbTab _
        & obj実行後データテーブル.get主キー項目連結文字列(lng変更前行番号) _
        & vbTab _
        & "変更後" & vbTab & lng変更後行番号 & "行目" & vbTab _
        & obj実行後データテーブル.get主キー項目連結文字列(lng変更後行番号)

End Sub

' *********************************************************************************************************************
' 機能：テーブルごとの座標情報を返却する
' *********************************************************************************************************************
'
Private Function getテーブル座標情報() As cls試験データテーブル()

    Dim lng最終行, lng最終列 As Long
    Dim varテーブル座標情報ワーク As Variant
    
    With ActiveSheet.UsedRange
        lng最終行 = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        lng最終列 = .Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    End With
    
    varテーブル座標情報ワーク = Range("A" & cnst試験データ開始行 & ":B" & lng最終行 + cnstヘッダ行数)
    
    Dim lngテーブル数 As Long
    lngテーブル数 = WorksheetFunction.CountA(Range("A" & cnst試験データ開始行 & ":A" & lng最終行))
    
    Dim obj試験データテーブル() As cls試験データテーブル
    ReDim obj試験データテーブル(1 To lngテーブル数)
    
    Dim i, index As Long: index = 0
    
    For i = 1 To UBound(varテーブル座標情報ワーク)
    
        ' テーブル名の列に値が設定されている場合
        If varテーブル座標情報ワーク(i, 1) <> "" Then
            
            index = index + 1
            
            Set obj試験データテーブル(index) = New cls試験データテーブル
            Call obj試験データテーブル(index).初期化(dic設定値, txtDBMS)
            
            ' 実際の行番号を設定するため開始行分行数を加算
            Set obj試験データテーブル(index).対象シート = ActiveSheet
            obj試験データテーブル(index).setテーブル名記載行 (cnst試験データ開始行 + i - 1)
            
            If index > 1 Then
            
                ' テーブル名が見つかった場合、その２セル上を前テーブルの終了行とみなす
                obj試験データテーブル(index - 1).setデータ終了行 (obj試験データテーブル(index).getテーブル名記載行 - 2)
            End If
        
        End If
        
    Next i
    
    If index > 0 Then
    
        ' 最後のテーブルは、最終行の3行後を最終行とみなす
        obj試験データテーブル(index).setデータ終了行 (lng最終行 + 3)
        
    End If
    
    getテーブル座標情報 = obj試験データテーブル
    
End Function

