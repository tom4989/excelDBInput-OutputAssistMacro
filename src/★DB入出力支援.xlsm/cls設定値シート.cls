VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls設定値シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'
Const cnstシート名 = "★設定値"
Const cnstブックシート名 = "*ブックシート名"

' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private objReg As Object

Private dic環境変数 As Object

Private dic設定値リスト As Object

Public Property Get 設定値リスト() As Object

    If dic設定値リスト Is Nothing Then
    
        Call MsgBox("cls設定値シート：設定値がロードされていません。", vbExclamation)
        End
    
    End If
    
    Set 設定値リスト = dic設定値リスト
End Property

' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

End Sub

' *********************************************************************************************************************
' 機能：指定されたシートの内容をロードし、Dictionaryオブジェクトに設定する。
' *********************************************************************************************************************
'
Public Sub ロード(Optional txt設定値シート As String = cnstシート名)

    ' 環境変数取得
    Call 環境変数取得

    ' 辞書初期化
    Set dic設定値リスト = CreateObject("Scripting.Dictionary")

    Dim varシートの値全体() As Variant
    varシートの値全体 = シート内容取得(ThisWorkbook.Sheets(txt設定値シート))
    
    Dim i As Long
    
    For i = LBound(varシートの値全体) To UBound(varシートの値全体)
            
        ' A列に★が設定されている行のみ対象
        If varシートの値全体(i, 1) = "★" Then
        
            Dim j As Long, lng開始列 As Long, lng最終列 As Long
            lng開始列 = -1
            lng最終列 = -1

            ' 要素数を数える
            For j = LBound(varシートの値全体, 2) + 1 To UBound(varシートの値全体, 2)
        
                If varシートの値全体(i, j) <> "" Then
                
                    If lng開始列 = -1 Then
                    
                        lng開始列 = j
                    
                    End If
                
                    lng最終列 = j
                End If
            Next j
    
            ' 設定値行が2階層以上設定されている場合
            If lng開始列 < lng最終列 Then
        
                Dim objDic As Object
                Set objDic = dic設定値リスト
        
                ' 1行に記載された情報を解析する
                For j = lng開始列 To lng最終列 - 1
        
                    ' 処理中の列が、最後の列でない場合
                    If j < lng最終列 - 1 Then
                
                        If objDic.Exists(varシートの値全体(i, j)) Then
                    
                            Set objDic = objDic.Item(varシートの値全体(i, j))
                    
                        Else
                            Dim objNewDic As Object
                            Set objNewDic = CreateObject("Scripting.Dictionary")
                            Call objDic.Add(varシートの値全体(i, j), objNewDic)
                        
                            Set objDic = objNewDic

                        End If
                    
                    ' 処理中の列が、最後の列の1つ前の場合
                    Else
                    
                        Dim var設定値() As Variant
                    
                        ' 既に登録済の場合は、文字列のリストに設定値を追加する
                        If objDic.Exists(varシートの値全体(i, j)) Then
                            
                            If varシートの値全体(i, j + 1) <> "---" Then
                            
                                var設定値 = objDic.Item(varシートの値全体(i, j))
                                ReDim Preserve var設定値(UBound(var設定値) + 1)
                                var設定値(UBound(var設定値)) = エスケープ(varシートの値全体(i, j), varシートの値全体(i, j + 1))
                    
                                objDic.Item(varシートの値全体(i, j)) = var設定値
                    
                            End If
                    
                        ' 次の行を確認し、同一階層の場合、設定値を文字列でなく文字列のリストにする
                        ElseIf 階層一致確認(varシートの値全体, i, j) Then
                        
                            ReDim var設定値(0)
                            var設定値(0) = エスケープ(varシートの値全体(i, j), varシートの値全体(i, j + 1))
                            Call objDic.Add(varシートの値全体(i, j), var設定値)
                            
                        Else
                    
                            Call objDic.Add(varシートの値全体(i, j), エスケープ(varシートの値全体(i, j), varシートの値全体(i, j + 1)))
                    
                        End If
                        
                        Exit For
                    
                    End If
                Next j
            End If
        End If
    Next i
    
End Sub

' *********************************************************************************************************************
' 機能：現在の設定値と、次の行の設定値が同一の項目か確認する
' *********************************************************************************************************************
'
Private Function 階層一致確認(ByRef var検証対象() As Variant, lng行数 As Long, lng列数 As Long) As Boolean

    ' 先の行がない場合、Falseを返す
    If lng行数 = UBound(var検証対象) Then
    
        階層一致確認 = False
        Exit Function
    End If

    Dim i, j As Long
    
    For j = 1 To lng列数
    
        ' 一致しない要素が存在した時点でFalse
        If var検証対象(lng行数, j) <> var検証対象(lng行数 + 1, j) Then
        
            階層一致確認 = False
            Exit Function
            
        End If
    Next j
    
    階層一致確認 = True

End Function

' *********************************************************************************************************************
' 機能：シートに記載されたブックシート名の[]を、Like演算子で使用できるようにエスケープする
' *********************************************************************************************************************
'
Function エスケープ(txt項目名, txt項目値) As String

    ' 環境変数の置換を行う
    If txt項目値 Like "*%*%*" Then
    
        Dim varKey As Variant
        
        For Each varKey In dic環境変数.Keys
        
            txt項目値 = Replace(txt項目値, "%" & varKey & "%", dic環境変数.Item(varKey))
            
        Next

    End If

    If Not (txt項目名 Like cnstブックシート名) Then
    
        エスケープ = txt項目値
        Exit Function
    End If

    If objReg Is Nothing Then
        Set objReg = CreateObject("VBScript.RegExp")
        
        objReg.Pattern = "^\["
    End If
    
    エスケープ = Replace(txt項目値, "]", "[]]")
    エスケープ = objReg.Replace(エスケープ, "[[]")
    
End Function

' *********************************************************************************************************************
' 機能：環境変数を取得する
' *********************************************************************************************************************
'
Private Function 環境変数取得()

    Set dic環境変数 = CreateObject("Scripting.Dictionary")

    Dim i As Long: i = 1

    Do
        Dim txtEnviron As String
        txtEnviron = Environ(i)
        
        Dim txtKey As String
        txtKey = UCase(Mid(txtEnviron, 1, InStr(txtEnviron, "=") - 1))

        Dim txtValue As String
        txtValue = Mid(txtEnviron, InStr(txtEnviron, "=") + 1)

        dic環境変数.Item(txtKey) = txtValue

        i = i + 1
    Loop Until Environ(i) = "" And i < 255

End Function

' *********************************************************************************************************************
' * 機能　：保存先が設定されている場合、保存して閉じる
' *********************************************************************************************************************
'
Public Sub 必要に応じて保存(wb対象ブック As Workbook, _
    Optional txt保存先キー As String = "保存先", Optional flgサイレントモード As Boolean = False)

    ' 保存先キーがない場合は何もしない
    If dic設定値リスト.Item(txt保存先キー) <> "" Then

        If wb対象ブック.Name Like "*.xlsx" Then
        
            wb対象ブック.Save
            Exit Sub
        
        End If

        Dim txt保存先ファイル As String
        txt保存先ファイル = f_日時サフィックス付与( _
            dic設定値リスト.Item(txt保存先キー), _
            dic設定値リスト.Item(txt保存先キー & "サフィックス"))

        ' ディレクトリがなかったら作成
        mkdirIFNotExist (funフォルダ名取得(txt保存先ファイル))

        wb対象ブック.SaveAs (txt保存先ファイル)

        ' サイレントモードの場合はファイルを閉じる
        If flgサイレントモード Then

            wb対象ブック.Close

        End If

    End If

End Sub

