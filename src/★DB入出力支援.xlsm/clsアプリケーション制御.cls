VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsアプリケーション制御"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   'エラートラップ
'   On Error GoTo ErrorCatch
'
'   objアプリケーション制御 as New clsアプリケーション制御
'   objアプリケーション制御.アプリケーション制御切替(False)
'
'   ' 本処理
'
'   ' 正常終了
'   GoTo Finally
'
' ErrorCatch:
'
' Finally:
'
'   '実行前の状態に戻す
'   objアプリケーション制御.アプリケーション制御切替(True)
'
' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
Private stateCalculation As XlCalculation

' *********************************************************************************************************************
' 機能：コンストラクタ
' *********************************************************************************************************************
'
Public Sub Class_Initialize()
    stateCalculation = True
End Sub

' *********************************************************************************************************************
' 機能：アプリケーション制御の切替
' *********************************************************************************************************************
'
Function アプリケーション制御切替(ByVal flg有効 As Boolean)

    ' １．画面更新制御
    Application.ScreenUpdating = flg有効
    
    ' ２．ステータスバー更新制御
    ' Application.DisplayStatusBar = flg有効
    
    ' ３．シートの再計算停止
    If Not flg有効 Then
        stateCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = stateCalculation
        Application.Calculate
    End If
        
    ' ４．イベント発生制御
    Application.EnableEvents = flg有効
    
    ' ５．プリンター通信無効化
    Application.PrintCommunication = flg有効

End Function


' *********************************************************************************************************************
' 機能：シート制御の切替
' *********************************************************************************************************************
'
Function シート制御切替(ByRef targetSheet As Worksheet, ByVal flg有効 As Boolean)

    ' ６．ページ区切り線非表示
    targetSheet.DisplayPageBreaks = flg有効
    
    ' ７．ピボットテーブル再計算停止
    targetSheet.DisplayPageBreaks = flg有効

End Function
