VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls設定シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------
'
Const cnstシート名 = "設定"
Const cnst設定項目 = "▼設定項目"
Const cnst日付関数 = "▼日付関数"
Const cnst日時関数 = "▼日時関数"
Const cnst折返文字数 = "▼折返文字数"
Const cnst結果ファイル出力先 As String = "▼結果ファイル出力先"
Const cnstテーブル物理名規則 = "▼結果ファイル出力先"
Const cnstテーブル論理名記載列 = "▼テーブル論理名記載列"
Const cnstテーブル物理名記載列 = "▼テーブル物理名記載列"
Const cnstテーブル件数記載列 = "▼テーブル件数記載列"
Const cnstカラム開始列 = "▼カラム開始列"

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
Public Enum DBMS区分
    Oralce = 1
    MySQL = 2
End Enum

Private txt日付関数 As String
Private txt日付形式 As String
Private txt日時関数 As String
Private txt日時形式 As String
Private lng折返文字数 As Long
Private txt結果ファイル出力先 As String
Private txtテーブル物理名規則 As String
Private txtテーブル論理名記載列 As String
Private txtテーブル物理名記載列 As String
Private txtテーブル件数記載列 As String
Private txtカラム開始列 As String

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------
'
Private enumDBMS区分値 As DBMS区分

Public Property Get DBMS区分値(argDBMS区分 As DBMS区分)
    DBMS区分 = enumDBMS区分
End Property

Public Property Get 日付関数() As String
    日付関数 = txt日付関数
End Property

Public Property Get 日付形式() As String
    日付形式 = txt日付形式
End Property

Public Property Get 日時関数() As String
    日時関数 = txt日時関数
End Property

Public Property Get 日時形式() As String
    日時形式 = txt日時形式
End Property

Public Property Get 折返文字数() As String
    折返文字数 = lng折返文字数
End Property

Public Property Get 結果ファイル出力先() As String
    結果ファイル出力先 = txt結果ファイル出力先
End Property

Public Property Get テーブル物理名規則() As String
    テーブル物理名規則 = txtテーブル物理名規則
End Property

Public Property Get テーブル論理名記載列() As String
    テーブル論理名記載列 = txtテーブル論理名記載列
End Property

Public Property Get テーブル物理名記載列() As String
    テーブル物理名記載列 = txtテーブル物理名記載列
End Property

Public Property Get テーブル件数記載列() As String
    テーブル件数記載列 = txtテーブル件数記載列
End Property

Public Property Get カラム開始列() As String
    カラム開始列 = txtカラム開始列
End Property

' *********************************************************************************************************************
' 機能：コンストラクタ
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

    ' 日付関数
    Dim var日付関数リスト As Variant
    Set var日付関数リスト = タイトル名指定でリスト値のRange情報を取得(cnst日付関数, ThisWorkbook.Sheets(cnstシート名))
    
    txt日付関数 = var日付関数リスト(1)
    txt日付形式 = var日付関数リスト(1).Offset(0, 1)
    
    ' 日時関数
    Dim var日時関数リスト As Variant
    Set var日時関数リスト = タイトル名指定でリスト値のRange情報を取得(cnst日時関数, ThisWorkbook.Sheets(cnstシート名))
    
    txt日時関数 = var日時関数リスト(1)
    txt日時形式 = var日時関数リスト(1).Offset(0, 1)
    
    ' 折返文字数
    Dim var折返文字数 As Variant
    Set var折返文字数 = タイトル名指定でリスト値のRange情報を取得(cnst折返文字数, ThisWorkbook.Sheets(cnstシート名))
    
    lng折返文字数 = CLng(var折返文字数(1))
    
    ' 結果ファイル出力先
    Dim var結果ファイル出力先 As Variant
    Set var結果ファイル出力先 = タイトル名指定でリスト値のRange情報を取得(cnst結果ファイル出力先, ThisWorkbook.Sheets(cnstシート名))
    
    txt結果ファイル出力先 = var結果ファイル出力先(1)
        
    ' テーブル論理名記載列
    log cnstテーブル論理名記載列
    
    Dim varテーブル論理名記載列 As Variant
    Set varテーブル論理名記載列 = タイトル名指定でリスト値のRange情報を取得(cnstテーブル論理名記載列, ThisWorkbook.Sheets(cnstシート名))
    
    txtテーブル論理名記載列 = varテーブル論理名記載列(1)

    ' テーブル物理名規則
    log cnstテーブル物理名規則
    
    Dim varテーブル物理名規則 As Variant
    Set varテーブル物理名規則 = タイトル名指定でリスト値のRange情報を取得(cnstテーブル物理名規則, ThisWorkbook.Sheets(cnstシート名))
    
    txtテーブル論理名記載列 = varテーブル物理名規則(1)

    ' テーブル物理名記載列
    Dim varテーブル物理名記載列 As Variant
    Set varテーブル物理名記載列 = タイトル名指定でリスト値のRange情報を取得(cnstテーブル物理名記載列, ThisWorkbook.Sheets(cnstシート名))
    
    txtテーブル物理名記載列 = varテーブル物理名記載列(1)

    ' テーブル件数記載列
    Dim varテーブル件数記載列 As Variant
    Set varテーブル件数記載列 = タイトル名指定でリスト値のRange情報を取得(cnstテーブル件数記載列, ThisWorkbook.Sheets(cnstシート名))
    
    txtテーブル件数記載列 = varテーブル件数記載列(1)

    ' カラム開始列
    Dim varカラム開始列 As Variant
    Set varカラム開始列 = タイトル名指定でリスト値のRange情報を取得(cnstカラム開始列, ThisWorkbook.Sheets(cnstシート名))
    
    txtカラム開始列 = varカラム開始列(1)
        
End Sub

' *********************************************************************************************************************
' 接続文字列の取得
' *********************************************************************************************************************
'
Public Function getConnectionString()

    ' 設定項目
    Dim obj設定値リスト As Variant
    Set obj設定値リスト = タイトル名指定でリスト値のRange情報を取得(cnst設定項目, ThisWorkbook.Sheets(cnstシート名))
    
    Dim txt接続文字列 As String

    Dim i As Long
    Dim obj設定値
    For Each obj設定値 In obj設定値リスト
    
        txt接続文字列 = txt接続文字列 & obj設定値 & "=" & obj設定値.Offset(0, 1) & ";"
        
        If obj設定値 Like "*MySQL*" Then
            enumDBMS区分値 = DBMS区分.MySQL
        End If
        
    Next obj設定値
    
    getConnectionString = txt接続文字列

End Function
