VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
Dim objStringBuilder As Object
Dim lngPrevLength As Long

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------
'
Property Get length() As Long
    length = objStringBuilder.length
End Property

Property Let length(length As Long)
    lngPrevLength = 0
    objStringBuilder.length = length
End Property

Property Get lastLineLength() As Long
    lastLineLength = objStringBuilder.length - lngPrevLength
End Property


' *********************************************************************************************************************
' 機能：コンストラクタ
' *********************************************************************************************************************
'
Private Sub Class_Initialize()
    lngPrevLength = 0
    Set objStringBuilder = CreateObject("System.Text.StringBuilder")
End Sub

' *********************************************************************************************************************
' 機能：デストラクタ
' *********************************************************************************************************************
'
Private Sub Class_Terminate()
    Set objStringBuilder = Nothing
End Sub

' *********************************************************************************************************************
' 機能：追加
' *********************************************************************************************************************
'
Public Sub append(str As String)
    Call objStringBuilder.Append_3(str)
End Sub

' *********************************************************************************************************************
' 機能：追加＋改行
' *********************************************************************************************************************
'
Public Sub appendLine(str As String)
    Call objStringBuilder.Append_3(str)
    Call objStringBuilder.Append_3(vbCrLf)
    
    lngPrevLength = objStringBuilder.length
End Sub

' *********************************************************************************************************************
' 機能：
' *********************************************************************************************************************
'
Public Sub insertHead(str As String)
    Call objStringBuilder.Insert_2(0, str)
End Sub

' *********************************************************************************************************************
' 機能：結果取得
' *********************************************************************************************************************
'
Public Function toString() As String
    toString = objStringBuilder.toString
End Function
