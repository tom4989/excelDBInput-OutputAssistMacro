VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------
'
' 文字列長
Private txtBuffer As String
Private lngBuferLength As Long
Private lngBufferIndex As Long

Private lngPrevLength As Long

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------
'
Property Get length() As Long
    length = lngBufferIndex
End Property

Property Let length(length As Long)
    
    lngPrevLength = 0
    lngBufferIndex = length

End Property

Property Get lastLineLength() As Long
    lastLineLength = lngBufferIndex - lngPrevLength
End Property


' *********************************************************************************************************************
' 機能：コンストラクタ
' *********************************************************************************************************************
'
Private Sub Class_Initialize()

    lngBuferLength = 0
    lngBufferIndex = 0
    lngPrevLength = 0

End Sub

' *********************************************************************************************************************
' 機能：デストラクタ
' *********************************************************************************************************************
'
Private Sub Class_Terminate()

End Sub

' *********************************************************************************************************************
' 機能：追加
' *********************************************************************************************************************
'
Public Sub append(txtAppend As String)

    ' 追加する文字列の長さを取得
    Dim lngAppendLength As Long
    lngAppendLength = Len(txtAppend)

    If lngAppendLength > 0 Then

        bufferCalculate (lngAppendLength)

        Mid$(txtBuffer, lngBufferIndex + 1, lngAppendLength) = txtAppend
        lngBufferIndex = lngBufferIndex + lngAppendLength
 
    End If
 
End Sub

' *********************************************************************************************************************
' 機能：追加＋改行
' *********************************************************************************************************************
'
Public Sub appendLine(str As String)
    Call append(str)
    Call append(vbCrLf)
    
    lngPrevLength = lngBuferLength
End Sub

' *********************************************************************************************************************
' 機能：
' *********************************************************************************************************************
'
Public Sub insertHead(txtAppend As String)

    ' 追加する文字列の長さを取得
    Dim lngAppendLength As Long
    lngAppendLength = Len(txtAppend)

    If lngAppendLength > 0 Then

        bufferCalculate (lngAppendLength)

        ' 元の文字を追加文字分スライド
        Mid$(txtBuffer, lngAppendLength + 1, lngBufferIndex) = Mid$(txtBuffer, 1, lngBufferIndex)

        ' 先頭にパラメータで指定された文字を追加
        Mid$(txtBuffer, 1, lngAppendLength) = txtAppend
    
        lngBufferIndex = lngBufferIndex + lngAppendLength

    End If

End Sub

' *********************************************************************************************************************
' 機能：バッファーの計算をする
' *********************************************************************************************************************
'
Private Sub bufferCalculate(lngAppendLength As Long)

    ' バッファが不足している場合
    If (lngBuferLength < lngBufferIndex + lngAppendLength) Then
    
        ' 追加するバッファを計算
        Dim lngAppendBufferLength As Long
        lngAppendBufferLength = (lngAppendLength \ 4096) * 4096 + 4096
        
        lngBuferLength = lngBuferLength + lngAppendBufferLength
        txtBuffer = txtBuffer & String(lngAppendBufferLength, Chr(0))
    End If

End Sub

' *********************************************************************************************************************
' 機能：結果取得
' *********************************************************************************************************************
'
Public Function toString() As String
    toString = Left$(txtBuffer, lngBufferIndex)
End Function
