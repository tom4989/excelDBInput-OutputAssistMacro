VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsテンプレート作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
' 共通の設定値
Private dic設定値 As Object

' テンプレート固有の設定値
Private dicテンプレート設定値 As Object

' テンプレートの開始位置
Private lngテンプレート開始位置 As Long

' テンプレートの行数
Private lngテンプレート行数 As Long

' DB接続オブジェクト
Private dbconn As Object

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

    Dim obj設定値シート As Object

    ' 共通の設定値の読込
    Set obj設定値シート = New cls設定値シート
    obj設定値シート.ロード
    
    Set dic設定値 = obj設定値シート.設定値リスト
    
    ' テンプレートの設定値の読込
    Set obj設定値シート = New cls設定値シート
    obj設定値シート.ロード (ActiveSheet.Name)
  
    Set dicテンプレート設定値 = obj設定値シート.設定値リスト
    
    lngテンプレート開始位置 = dicテンプレート設定値.Item("テンプレート開始位置")
    lngテンプレート行数 = dicテンプレート設定値.Item("テンプレート行数")

    ' DB接続
    Set dbconn = CreateObject("ADODB.Connection")

    dbconn.ConnectionString = f_Dictonary結合( _
        dic設定値("接続情報").Item(dicテンプレート設定値.Item("接続情報")))

    dbconn.Open

End Sub

' *********************************************************************************************************************
' 機能：終了時の処理
' *********************************************************************************************************************
'
Public Sub Class_Terminate()

    dbconn.Close

    Set dbconn = Nothing
End Sub

' *********************************************************************************************************************
' 機能：テンプレートの作成
' *********************************************************************************************************************
'
Public Sub テンプレート作成()

    ' 出力対象のテンプレート分処理を繰り返す
    Dim txtテンプレート名 As Variant
    
    For Each txtテンプレート名 In dicテンプレート設定値.Item("テンプレート").Keys
    
        Call f_シートコピー(ThisWorkbook, "★テンプレート書式", CStr(txtテンプレート名))
        
        Dim lng現在位置 As Long
        lng現在位置 = lngテンプレート開始位置
        
        Dim txtテーブル名 As Variant
        
        For Each txtテーブル名 In dicテンプレート設定値.Item("テンプレート").Item(txtテンプレート名)
        
            Call テーブル作成(ThisWorkbook.Sheets(txtテンプレート名), CStr(txtテーブル名), lng現在位置)
            lng現在位置 = lng現在位置 + lngテンプレート行数
            
        Next txtテーブル名
        
        Call 不要行の削除( _
            ThisWorkbook.Sheets(txtテンプレート名), lngテンプレート開始位置)
        
    Next txtテンプレート名

End Sub

' *********************************************************************************************************************
' 機能：テンプレートのテーブル部分の作成
' *********************************************************************************************************************
'
Private Sub テーブル作成(ws対象シート As Worksheet, txtテーブル名 As String, index As Long)

    Dim lngEnd As Long
    lngEnd = index + lngテンプレート行数

    With ws対象シート
    
        .Rows(lngテンプレート開始位置 & ":" & lngテンプレート開始位置 + lngテンプレート行数 - 1).Select
        Selection.Copy
        .Rows(lngEnd & ":" & lngEnd).Select
        Selection.Insert Shift:=xlDown

        Dim txtSQL As String
        txtSQL = Replace(dicテンプレート設定値.Item("メタデータ取得SQL"), ":テーブル名", "'" & txtテーブル名 & "'")

        Debug.Print (txtSQL)

        Dim rs As Object
        Set rs = dbconn.Execute(txtSQL)

        .Cells(lngEnd, 1) = txtテーブル名
        .Cells(lngEnd, 4) = rs.Fields("TABLE_COMMENTS")
        .Cells(lngEnd, 4).WrapText = False

        Dim j As Long: j = 2
        Do Until rs.EOF
        
            ' 桁情報の編集
            Dim txt桁 As String: txt桁 = ""
            
            If f_配列含まれているかチェック( _
                dicテンプレート設定値.Item("データ型桁数").Item("DATA_LENGTH"), rs.Fields("DATA_TYPE")) Then
                
                txt桁 = "(" & rs.Fields("DATA_LENGTH")
                
            End If

            If f_配列含まれているかチェック( _
                dicテンプレート設定値.Item("データ型桁数").Item("DATA_PRECISION"), rs.Fields("DATA_TYPE")) Then
                
                If rs.Fields("DATA_PRECISION") <> "" Then
                
                    txt桁 = "(" & rs.Fields("DATA_PRECISION")

                End If

            End If

            If f_配列含まれているかチェック( _
                dicテンプレート設定値.Item("データ型桁数").Item("DATA_SCALE"), rs.Fields("DATA_TYPE")) _
                And rs.Fields("DATA_SCALE") <> "" And rs.Fields("DATA_SCALE") <> 0 Then

                If txt桁 <> "" Then
                
                    txt桁 = txt桁 & ", "
                
                End If
                
                txt桁 = txt桁 & rs.Fields("DATA_SCALE")

            End If

            If txt桁 <> "" Then
            
                txt桁 = txt桁 & ")"
            
            End If

            ' PrimaryKey/NotNull情報の編集
            Dim txtPK As String: txtPK = ""
            
            If rs.Fields("POSITION") <> "" Then
            
                txtPK = rs.Fields("POSITION")
                
            ElseIf rs.Fields("NULLABLE") = "N" Then
            
                txtPK = "〇"
            
            End If

            ' 取得、編集したメタデータをテンプレートに反映
            .Cells(lngEnd + 1, j) = rs.Fields("COLUMN_NAME")
            .Cells(lngEnd + 2, j) = rs.Fields("COL_COMMENTS")
            .Cells(lngEnd + 3, j) = rs.Fields("DATA_TYPE") & txt桁
            .Cells(lngEnd + 4, j) = txtPK

            j = j + 1
            rs.MoveNext
        Loop

        ' 書式のコピー
        .Range( _
            .Cells(lngテンプレート開始位置 + 1, 2), _
            .Cells(lngテンプレート開始位置 + lngテンプレート行数 - 2, 2) _
        ).Select
        Selection.Copy

        .Range( _
            .Cells(lngEnd + 1, 2), _
            .Cells(lngEnd + lngテンプレート行数 - 2, j - 1) _
        ).Select
        Selection.PasteSpecial _
            Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    End With
    
End Sub

' *********************************************************************************************************************
' 機能：テンプレートの不要行の削除
' *********************************************************************************************************************
'
Private Sub 不要行の削除(ws対象シート As Worksheet, index As Long)

    Dim lngEnd As Long
    lngEnd = index + lngテンプレート行数

    With ws対象シート
    
        .Rows(index & ":" & lngEnd - 1).Select
        Selection.Delete

    End With
    
End Sub
