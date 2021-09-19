VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls試験データテーブル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

Const cnstインデント = "   "

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

' 対象シート
Private obj対象シート As Worksheet

Public Enum 型種別
    NUMBER系 = 1
    TIMESTAMP系 = 2
    DATE系 = 3
    CHAR系 = 4
End Enum

Private Type 型桁情報
    型名 As String
    型種別 As 型種別
    桁数 As Long
End Type

' 行情報
Private lngテーブル名記載行 As Long
Private lngカラム論理名記載行 As Long
Private lngカラム物理名記載行 As Long
Private lng型桁記載行 As Long
Private lng制約記載行 As Long
Private lngデータ開始行 As Long
Private lngデータ終了行 As Long

' 列情報
Private lngカラム開始列 As Long
Private lngカラム終了列 As Long

' 名称
Private txtテーブル論理名, txtテーブル物理名 As String

' 状態
Private isHidden As Boolean
Private lngDBCount結果 As Long

Private dic設定値 As Object

Private txtDBMS As String

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------

Public Property Set 対象シート(ByRef arg対象シート)
    Set obj対象シート = arg対象シート
End Property

Public Property Get 対象シート() As Worksheet

    If obj対象シート Is Nothing Then
        Set obj対象シート = ActiveSheet
    End If
    
    Set 対象シート = obj対象シート
    
End Property

Public Property Get カラム論理名記載行()
    カラム論理名記載行 = lngカラム論理名記載行
End Property

Public Property Get データ開始行()
    データ開始行 = lngデータ開始行
End Property

Public Property Get データ終了行()
    データ終了行 = lngデータ終了行
End Property

Public Property Get カラム終了列()
    カラム終了列 = lngカラム終了列
End Property


' *********************************************************************************************************************
' 機能：初期化
' *********************************************************************************************************************
'
Public Sub 初期化(arg設定値 As Object, argDBMS As String, arg対象シート As Worksheet)

    Set dic設定値 = arg設定値
    Set obj対象シート = arg対象シート

    ' DBMSの種類を設定
    txtDBMS = argDBMS

End Sub

' *********************************************************************************************************************
' 機能：テーブル名記載行（＝試験データテーブルの開始位置）を設定する。
' *********************************************************************************************************************
'
Public Sub setテーブル名記載行(ByVal argテーブル名記載行)

    ' 行情報の設定
    lngテーブル名記載行 = argテーブル名記載行
    lngカラム物理名記載行 = argテーブル名記載行 + 1
    lngカラム論理名記載行 = argテーブル名記載行 + 2
    lng型桁記載行 = argテーブル名記載行 + 3
    lng制約記載行 = argテーブル名記載行 + 4
    lngデータ開始行 = argテーブル名記載行 + 5
    
    ' 列情報の設定
    lngカラム開始列 = CAlp2Num(dic設定値.Item("カラム開始列"))
    
    If obj対象シート.Cells(lngカラム物理名記載行, lngカラム開始列 + 1) <> "" Then
    
        lngカラム終了列 = obj対象シート.Cells(lngカラム物理名記載行, lngカラム開始列).End(xlToRight).Column
    Else
        lngカラム終了列 = lngカラム開始列
        
    End If

    ' 名称情報の設定
    txtテーブル論理名 = obj対象シート.Range("D" & lngテーブル名記載行).Value
    txtテーブル物理名 = obj対象シート.Range("A" & lngテーブル名記載行).Value

    ' 表示/非表示状態(論理名列の状態で判断)
    isHidden = obj対象シート.Cells(lngカラム論理名記載行, 1).EntireRow.Hidden

End Sub


' *********************************************************************************************************************
' 機能：テーブル名記載行（＝試験データテーブルの開始位置）を返却する。
' *********************************************************************************************************************
'
Public Function getテーブル名記載行() As Long

    getテーブル名記載行 = lngテーブル名記載行

End Function


' *********************************************************************************************************************
' 機能：データ終了位置を設定する。本項目については、他テーブル
' *********************************************************************************************************************
'
Public Sub setデータ終了行(ByVal param As Long)

    lngデータ終了行 = param
    
End Sub


' *********************************************************************************************************************
' 機能：非表示かどうかを示す
' *********************************************************************************************************************
'
Public Function Hidden()

    Hidden = isHidden
    
End Function


' *********************************************************************************************************************
' 機能：テーブル論理名を返却する。
' *********************************************************************************************************************
'
Public Function getテーブル物理名()

    getテーブル物理名 = txtテーブル物理名

End Function

' *********************************************************************************************************************
' 機能：テーブル物理名を返却する。
' *********************************************************************************************************************
'
Public Function getテーブル論理名()

    getテーブル論理名 = txtテーブル論理名

End Function

' *********************************************************************************************************************
' 機能：指定された列番号のカラムの論理名を返却する。
' *********************************************************************************************************************
'
Public Function getカラム論理名(ByVal arg指定カラム列 As Long) As String

    getカラム論理名 = Me.対象シート.Cells(lngカラム論理名記載行, arg指定カラム列)

End Function

' *********************************************************************************************************************
' 機能：DBカウント結果を返却する。
' *********************************************************************************************************************
'
Public Function setDBCount結果(ByVal param As Long)

    lngDBCount結果 = param
    
End Function

' *********************************************************************************************************************
' 機能：DBカウント結果を設定する。
' *********************************************************************************************************************
'
Public Function getDBCount結果() As Long

    getDBCount結果 = lngDBCount結果
    
End Function


' *********************************************************************************************************************
' 機能：セルに入力されているデータ件数を返却する。
' *********************************************************************************************************************
'
Public Function get件数()

    get件数 = Me.対象シート.Cells(lngテーブル名記載行, 7)

End Function

' *********************************************************************************************************************
' 機能：SELECT文に、ORDER BYを付与する。
' *********************************************************************************************************************
'
Public Function addOrderBy(ByVal txtQuery As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "(.+? FROM)"
    
    If txtQuery Like "* UNION *" Then
    
        addOrderBy = Re.Replace(txtQuery, "SELECT * FROM ( $1") & " ) DUMMY "
    
    Else
        addOrderBy = txtQuery
        
    End If

    Dim var主キー As Variant
    var主キー = get主キー()
    
    If Not IsEmpty(var主キー) Then
        addOrderBy = addOrderBy & " ORDER BY " & Join(var主キー, ", ")
    End If

End Function

' *********************************************************************************************************************
' 機能：1行のデータ行をもとにSQL文を作成する。
' 　　　データ行を指定しない場合、WHERE句なしのSELECT文を作成する。
' *********************************************************************************************************************
'
Public Function createSELECT文From単行(Optional argデータ行 As Long = -1) As String

    With ActiveSheet
        
        Dim j As Long
        Dim stbSELECT文 As New clsStringBuilder
        Dim stbWHERE句 As New clsStringBuilder
        
        For j = 2 To lngカラム終了列
        
            ' ---------------------------------------------------------------------------------------------------------
            ' SELECT句
            ' ---------------------------------------------------------------------------------------------------------
        
            If stbSELECT文.length = 0 Then
                stbSELECT文.append ("SELECT ")
            Else
                If stbSELECT文.lastLineLength > dic設定値.Item("折返文字数") Then
                    Call stbSELECT文.appendLine("")
                End If
            
                stbSELECT文.append (", ")
            End If
        
            Dim txtカラム値 As String
            txtカラム値 = editカラム値(.Cells(lngカラム物理名記載行, j).Value, .Cells(lng型桁記載行, j).Value, True)

            If txtカラム値 Like "*(*" Then
                ' 列名に関数を使用している場合、別名を付与する
                stbSELECT文.append (txtカラム値 & " " & .Cells(lngカラム物理名記載行, j).Value)
            Else
                stbSELECT文.append (txtカラム値)
            End If
            
            ' ---------------------------------------------------------------------------------------------------------
            ' WHERE句
            ' ---------------------------------------------------------------------------------------------------------
        
            If argデータ行 = -1 Then
                GoTo continue
            End If
        
            If .Cells(argデータ行, j).Value <> "" Then
        
                If stbWHERE句.length = 0 Then
                    stbWHERE句.append (" WHERE ")
                Else
                    If stbWHERE句.lastLineLength > dic設定値.Item("折返文字数") Then
                        Call stbWHERE句.appendLine("")
                    End If
                    stbWHERE句.append (" AND ")
                End If
            
                stbWHERE句.append (.Cells(lngカラム物理名記載行, j).Value)
                stbWHERE句.append (" = ")
                stbWHERE句.append (editカラム値(.Cells(argデータ行, j).Value, .Cells(lng型桁記載行, j).Value, False))
                
            End If
continue:
        Next j
        
    End With

    createSELECT文From単行 = stbSELECT文.toString & " FROM " & txtテーブル物理名 & stbWHERE句.toString

End Function

' *********************************************************************************************************************
' 機能：複数のデータ行をもとにSQL文を作成する。
' 　　　作成するSELECT文は、1テーブル（複数行）に対し、1SELECT文（複数のSELECT文をUNIONでまとめたもの）となる。
' *********************************************************************************************************************
'
Public Function createSELECT文From複数行() As String

    Dim stbQuery As New clsStringBuilder

    If getデータ行の入力数() > 0 Then
    
        Dim j As Long
        
        For j = lngデータ開始行 To lngデータ終了行
        
            If getデータ行の入力数(j) > 0 Then
            
                If stbQuery.length <> 0 Then
                    
                    stbQuery.append (vbCrLf & " UNION ")
                End If
            
                stbQuery.append (createSELECT文From単行(j))
            End If
             
         Next j
         
    Else
    
    stbQuery.append (createSELECT文From単行())
        
    End If
        
    ' -----------------------------------------------------------------------------------------------------------------
    ' ORDER B句の付与
    ' -----------------------------------------------------------------------------------------------------------------
        
    createSELECT文From複数行 = addOrderBy(stbQuery.toString)
        
End Function


' *********************************************************************************************************************
' 機能：複数のデータ行を元にSQL文を作成する。作成するSELECT文は、1データに対し、1SELECT文となる。
' *********************************************************************************************************************
'
Public Function createSELECT文From複数行To複数SQL(Optional ByVal is選択行のみ As Boolean = False) As String

    Dim txtQuery As String
    
    ' データ行のいずれかに何かしら入力されている場合
    If getデータ行の入力数() > 0 Then
    
        Dim j As Long

        For j = lngデータ開始行 To lngデータ終了行
        
            If is選択行のみ And Not is選択状態(j) Then ' 選択行のみSQL文作成対象にする場合の考慮
            
                GoTo jContinue
            End If
            
            If getデータ行の入力数(j) > 0 Then
            
                txtQuery = txtQuery & addOrderBy(createSELECT文From単行(j)) & ";" & vbCrLf
            
            End If
jContinue:

        Next j

    End If

    ' 前処理でSQLが作成されていない場合
    
    If txtQuery = "" Then
    
        Dim k As Long
        
        For k = lngデータ開始行 To lngデータ終了行
        
            If is選択行のみ And Not is選択状態(k) Then ' 選択行のみSQL文作成対象にする場合の考慮
                ' 何もしない
            Else
            
                txtQuery = addOrderBy(createSELECT文From単行()) & ";" & vbCrLf
                Exit For ' 複数行選択されていても、同じSQLになるので1行だけ作成する
            End If
        Next k
    End If

    If txtQuery <> "" Then
    
        txtQuery = vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル論理名 & vbCrLf & txtQuery
        createSELECT文From複数行To複数SQL = txtQuery
        
    End If
        
End Function

' *********************************************************************************************************************
' 機能：INSERT文を作成する
' *********************************************************************************************************************
'
Public Function createInsert文(ByRef obj引数対象シート As Worksheet, Optional ByVal is選択行のみ As Boolean = False) As String

    Dim txt結果 As New clsStringBuilder
    Dim txtInsertInto As New clsStringBuilder
    Dim txtInsertValues As New clsStringBuilder
    
    With obj引数対象シート
    
        If .Range("B" & lngデータ開始行).Value = "" Then
        
            Exit Function ' データがない場合、INSERT文の作成はしない
        End If
        
        Dim j, k As Long
        
        ' データ行、1行ごとの処理
        For j = lngデータ開始行 To lngデータ終了行
        
            If .Range("B" & j).Value = "" Then
            
                GoTo jContinue
            End If
            
            If is選択行のみ And Not is選択状態(j) Then ' 選択行のみSQL文作成対象にする場合の考慮
            
                GoTo jContinue
            End If

            Call txtInsertInto.append("INSERT INTO " & txtテーブル物理名 & " (")
            Call txtInsertValues.append(" VALUES (")
            
            For k = 2 To lngカラム終了列
            
                If k > 2 Then
                    If txtInsertInto.lastLineLength > dic設定値.Item("折返文字数") Then
                        Call txtInsertInto.appendLine("")
                    End If
                    Call txtInsertInto.append(", ")
                    
                    If txtInsertValues.lastLineLength > dic設定値.Item("折返文字数") Then
                        Call txtInsertValues.appendLine("")
                    End If
                    Call txtInsertValues.append(", ")
                End If
                
                Call txtInsertInto.append(.Cells(lngカラム物理名記載行, k))
                Call txtInsertValues.append(editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k)))
            Next k
            
            Call txt結果.append(txtInsertInto.toString)
            Call txt結果.append(")")
            Call txt結果.append(vbCrLf)
            Call txt結果.append("   ")
            Call txt結果.append(txtInsertValues.toString)
            Call txt結果.append(");")
            Call txt結果.append(vbCrLf)
            
            txtInsertInto.length = 0
            txtInsertValues.length = 0
jContinue:
        Next j
        
    End With

    If txt結果.length <> 0 Then
        
        Call txt結果.insertHead(vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル物理名 & vbCrLf)
        
        createInsert文 = txt結果.toString
        
    End If

    Set txt結果 = Nothing

End Function

' *********************************************************************************************************************
' 機能：UPDATE文を作成する
' *********************************************************************************************************************
'
Public Function createUpdate文(ByRef obj引数対象シート As Worksheet, Optional ByVal is選択行のみ As Boolean) As String

    Dim stb結果 As New clsStringBuilder

    Dim stbUpdate As New clsStringBuilder
    Dim stbWHERE As New clsStringBuilder
                
    With Me.対象シート
    
        If .Range("B" & lngデータ開始行).Value = "" Then
        
            Exit Function ' データがない場合、UPDATE文の作成はしない
        End If
        
        Dim var主キー As Variant
        var主キー = Me.get主キー()
        
        Dim j, k As Long
        
        ' データ行、1行ごとの処理
        For j = lngデータ開始行 To lngデータ終了行
        
            If .Range("B" & j).Value = "" Then
            
                GoTo jContinue
            End If
            
            If is選択行のみ And Not is選択状態(j) Then ' 選択行のみSQL文作成対象にする場合の考慮
            
                GoTo jContinue
            End If
            
            For k = 2 To lngカラム終了列
            
                If containArray(var主キー, .Cells(lngカラム物理名記載行, k)) Then
                
                    If stbWHERE.length = 0 Then
                    
                        stbWHERE.append (" WHERE ")
                    Else
                        If stbWHERE.lastLineLength > dic設定値.Item("折返文字数") Then
                            Call stbWHERE.appendLine("")
                        End If
                        stbWHERE.append (" AND ")
                    
                    End If

                    stbWHERE.append (.Cells(lngカラム物理名記載行, k))
                    stbWHERE.append (" = ")
                    stbWHERE.append (editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k)))
                Else
                    If stbUpdate.length = 0 Then
                    
                        stbUpdate.append ("UPDATE " & txtテーブル物理名 & " SET ")
                    Else
                        If stbUpdate.lastLineLength > dic設定値.Item("折返文字数") Then
                            Call stbUpdate.appendLine("")
                        End If
                        stbUpdate.append (" , ")
                    End If
                    
                    stbUpdate.append (.Cells(lngカラム物理名記載行, k))
                    stbUpdate.append (" = ")
                    stbUpdate.append (editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k)))
                End If
                
            Next k
            
            stb結果.append (stbUpdate.toString & vbCrLf)
            stb結果.append ("   ")
            stb結果.append (stbWHERE.toString & ";" & vbCrLf)
            
            stbUpdate.length = 0
            stbWHERE.length = 0
jContinue:
        Next j
        
    End With

    If stb結果.length <> 0 Then
        stb結果.insertHead (vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル物理名 & vbCrLf)
        createUpdate文 = stb結果.toString
    End If

    Set stb結果 = Nothing
    Set stbUpdate = Nothing
    Set stbWHERE = Nothing

End Function

' *********************************************************************************************************************
' 機能：DELETE文を作成する
' *********************************************************************************************************************
'
Public Function createDelete文(ByRef obj引数対象シート As Worksheet, Optional ByVal is選択行のみ As Boolean) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Global = True
    Re.Pattern = "UPDATE [^;]+\n   "

    Dim txt結果 As String
    txt結果 = Re.Replace(createUpdate文(obj引数対象シート, is選択行のみ), "DELETE FROM " & txtテーブル物理名)

    createDelete文 = txt結果

End Function

' *********************************************************************************************************************
' 機能：SELECT文を、COUNTを行うSQLに変更する。
' *********************************************************************************************************************
'
Public Function createCount文(ByVal txtQuery As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "(SELECT .+? FROM)"
    
    createCount文 = Re.Replace(Replace(txtQuery, vbCrLf, ""), "SELECT COUNT(*) AS COUNT FROM")
    
End Function

' *********************************************************************************************************************
' 機能：枠をクリップボードにコピーする
' *********************************************************************************************************************
'
Public Sub copy枠Toクリップボード()

    ActiveWorkbook.ActiveSheet.Rows(lngテーブル名記載行 & ":" & lngデータ開始行).Copy
    
End Sub


' *********************************************************************************************************************
' 機能：データ行の内容をクリアする
' *********************************************************************************************************************
'
Public Sub clearデータ行()

    Me.対象シート.Rows(lngデータ開始行 & ";" & lngデータ終了行).ClearContents ' 内容をクリア
    Me.対象シート.Rows(lngデータ開始行 & ";" & lngデータ終了行).ClearComments ' コメントをクリア
    
End Sub

' *********************************************************************************************************************
' 機能：引数で指定された自動設定値と同一のカラム名がテーブル内に存在する場合、設定値を1行目にセットする。
' *********************************************************************************************************************
'
Public Sub set抽出条件(ByVal var自動設定値 As Variant)

    With Me.対象シート
    
        Dim i As Long
        For i = 2 To lngカラム終了列
    
            Dim j As Long
            For j = LBound(var自動設定値) To UBound(var自動設定値)
        
                If .Cells(lngカラム物理名記載行, i) = var自動設定値(j, 1) Then
            
                    .Cells(lngカラム物理名記載行, i) = var自動設定値(j, 2)
                
                End If
            Next j
        Next i
    End With
End Sub

' *********************************************************************************************************************
' 機能：指定された行に空行を作成する
' *********************************************************************************************************************
'
Public Sub add空行(ByVal arg追加行番号 As Long)

    Me.対象シート.Rows(lngデータ開始行).Copy  ' データ行の1行目から書式コピー
    Me.対象シート.Rows(arg追加行番号).Insert ' 行追加
    Me.対象シート.Rows(arg追加行番号).ClearContents ' 内容をクリア
    Me.対象シート.Rows(arg追加行番号).ClearComments ' コメントをクリア
    
    Application.CutCopyMode = False
    
End Sub


' *********************************************************************************************************************
' 機能：選択された行を黄色で着色
' *********************************************************************************************************************
'
Public Sub edit選択行強調(ByVal arg選択行番号 As Long)

    With Me.対象シート
    
        Call Me.edit変更強調色(.Range(.Cells(arg選択行番号, 2), .Cells(arg選択行番号, lngカラム終了列)))

    End With

End Sub

' *********************************************************************************************************************
' 機能：選択された行に色を付けて強調
' *********************************************************************************************************************
'
Public Sub edit変更強調色(ByRef arg修飾範囲 As Range)

    With arg修飾範囲.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub


' *********************************************************************************************************************
' 機能：選択された行を網掛
' *********************************************************************************************************************
'
Public Sub edit選択行網掛(ByVal arg選択行番号 As Long)

    With Me.対象シート
    
        Call Me.edit網掛(.Range(.Cells(arg選択行番号, 2), .Cells(arg選択行番号, lngカラム終了列)))
        
    End With

End Sub

' *********************************************************************************************************************
' 機能：選択された行に網掛け
' *********************************************************************************************************************
'
Public Sub edit網掛(ByRef rng修飾範囲 As Range)

    With rng修飾範囲.Interior
        .Pattern = xlGray16
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub


' *********************************************************************************************************************
' 機能：指定された主キー項目を連結した文字列を返却する
' *********************************************************************************************************************
'
Function get主キー項目連結文字列(ByVal arg対象データ行 As Long) As String

    get主キー項目連結文字列 = Join(get主キー(arg対象データ行))

End Function

' *********************************************************************************************************************
' 機能：データ行の入力数を返却する。
' *********************************************************************************************************************
'
Private Function getデータ行の入力数(Optional arg対象データ行 = -1)

    With ActiveSheet
    
        If arg対象データ行 = -1 Then
        
            getデータ行の入力数 = WorksheetFunction.CountA( _
                .Range(.Cells(lngデータ開始行, 2), .Cells(lngデータ終了行, lngカラム終了列)))
            
        Else
            
            getデータ行の入力数 = WorksheetFunction.CountA( _
                .Range(.Cells(arg対象データ行, 2), .Cells(arg対象データ行, lngカラム終了列)))
                
        End If
            
    End With

End Function

' *********************************************************************************************************************
' 機能：主キーを配列で返却する
' *********************************************************************************************************************
'
Public Function get主キー(Optional ByVal arg対象データ行 As Long = -1) As Variant

    If arg対象データ行 = -1 Then
        arg対象データ行 = lngカラム物理名記載行
    End If

    Dim var主キー As Variant
    ReDim var主キー(1 To lngカラム終了列) ' 予想され得る最大値＝カラム数で配列を確保
    
    Dim i, lng主キー数 As Long
    
    For i = 2 To lngカラム終了列
    
        With Me.対象シート

            If IsNumeric(.Cells(lng制約記載行, i).Value) And .Cells(lng制約記載行, i).Value <> "" Then
                
                lng主キー数 = lng主キー数 + 1
                var主キー(lng主キー数) = .Cells(arg対象データ行, i).Value
                
            End If
        End With
    Next
    
    ' 主キーがない場合、全列主キーとみなす
    If lng主キー数 = 0 Then
    
        For i = 2 To lngカラム終了列
    
            With Me.対象シート

                lng主キー数 = lng主キー数 + 1
                var主キー(lng主キー数) = .Cells(arg対象データ行, i).Value
                
            End With
        Next
    End If
    
    ReDim Preserve var主キー(1 To lng主キー数)
    get主キー = var主キー
    
End Function

' *********************************************************************************************************************
' 機能：LOB列のファイル名を主キーを含めた形式に変更する
' *********************************************************************************************************************
'
Public Sub LOB出力ファイルパス編集()

    Dim i As Long
    
    If lngデータ終了行 = 0 Then
    
        lngデータ終了行 = 最終行取得(obj対象シート, True)

    End If
    
    With obj対象シート
        For i = lngカラム開始列 To lngカラム終了列
    
            If .Cells(lng型桁記載行, i) Like "*LOB*" Then
            
                Dim j As Long
                For j = lngデータ開始行 To lngデータ終了行
                
                    If .Cells(j, i) <> "" Then
                    
                        Dim txtリネーム前ファイル名 As String
                        txtリネーム前ファイル名 = obj対象シート.Parent.Path & "\" & .Cells(j, i)
                    
                        Dim txtリネーム後ファイル名 As String
                        txtリネーム後ファイル名 = funフォルダ名取得(txtリネーム前ファイル名) & "\" & txtテーブル物理名 & "_" & get主キーと値を結合した文字列(j)
                        Name txtリネーム前ファイル名 As txtリネーム後ファイル名
                        
                        Dim txtセルに記載するパス As String
                        txtセルに記載するパス = Replace(txtリネーム後ファイル名, obj対象シート.Parent.Path & "\", "")
                        .Cells(j, i) = txtセルに記載するパス
                        
                        Call .Cells(j, i).Hyperlinks.Add(.Cells(j, i), Address:=funフォルダ名取得(txtセルに記載するパス))
                        
                    End If
                
                Next j
            End If
        Next i
    End With

End Sub

' *********************************************************************************************************************
' 機能：主キーと値を結合した文字列を返却する
' *********************************************************************************************************************
'
Public Function get主キーと値を結合した文字列(arg対象データ行 As Long) As String

    Dim txt返却結果 As String
    ReDim var主キー(1 To lngカラム終了列) ' 予想され得る最大値＝カラム数で配列を確保
    
    Dim i, lng主キー数 As Long
    
    For i = 2 To lngカラム終了列
    
        With Me.対象シート

            If IsNumeric(.Cells(lng制約記載行, i).Value) And .Cells(lng制約記載行, i).Value <> "" Then
                
                If txt返却結果 <> "" Then
                
                    txt返却結果 = txt返却結果 & "_"
                End If
                
                txt返却結果 = txt返却結果 & _
                    .Cells(lngカラム物理名記載行, i).Value & "-" & .Cells(arg対象データ行, i).Value
                
            End If
        End With
    Next
    
    ' 主キーがない場合、全列主キーとみなす
    If txt返却結果 = "" Then
    
        For i = 2 To lngカラム終了列
    
            With Me.対象シート

                If Not .Cells(lng型桁記載行, i) Like "*LOB*" Then
                
                    If txt返却結果 <> "" Then
                
                        txt返却結果 = txt返却結果 & "_"
                    End If
                
                    txt返却結果 = txt返却結果 & _
                        .Cells(lngカラム物理名記載行, i).Value & "-" & .Cells(arg対象データ行, i).Value
                End If
                
            End With
        Next
    End If
    
    get主キーと値を結合した文字列 = txt返却結果
    
End Function

' *********************************************************************************************************************
' 機能：カラムに対する値を型桁に合わせて加工する(チェック機能付)
' *********************************************************************************************************************
'
Private Function editカラム値( _
    ByVal argカラム値 As String, ByVal arg型桁 As String, Optional ByVal is列名 = False) As String
    
    Dim dicDBMS As Object
    Set dicDBMS = dic設定値.Item(dic設定値.Item("接続情報").Item(txtDBMS).Item("DBMS"))
    
    arg型桁 = UCase(arg型桁)
    
    If argカラム値 = "" Or argカラム値 Like "*(NULL)*" Then
        editカラム値 = "NULL"
        Exit Function
    End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    If arg型桁 Like "DATE*" Then
    
        If UCase(argカラム値) = "SYSTIMESTAMP" Or UCase(argカラム値) = "SYSDATE" Then
        
            editカラム値 = argカラム値
        Else
            If is列名 Then
                editカラム値 = dicDBMS.Item("文字列関数") & "(" & argカラム値 & ", '" & dicDBMS.Item("日付形式") & "')"
            Else
                editカラム値 = dicDBMS.Item("日付関数") & "('" & argカラム値 & "', '" & dicDBMS.Item("日付形式") & "')"
            End If
        End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg型桁 Like "TIMESTAMP*" Then
        
        If UCase(argカラム値) = "SYSTIMESTAMP" Or UCase(argカラム値) = "SYSDATE" Then
        
            editカラム値 = argカラム値
        Else
            If is列名 Then
                editカラム値 = dicDBMS.Item("文字列関数") & "(" & argカラム値 & ", '" & dicDBMS.Item("日時形式") & "')"
            Else
                editカラム値 = dicDBMS.Item("日時関数") & "('" & argカラム値 & "', '" & dicDBMS.Item("日時形式") & "')"
            End If
        End If

    ' -----------------------------------------------------------------------------------------------------------------
    ' NUMBER、INT、FLOAT型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg型桁 Like "NUMBER*" Or arg型桁 Like "INT*" Or arg型桁 Like "FLOAT*" Then
    
        editカラム値 = argカラム値

    ' -----------------------------------------------------------------------------------------------------------------
    ' VARCHAR2,CHAR型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg型桁 Like "VARCHAR*" Or arg型桁 Like "CHAR*" Or arg型桁 Like "BLOB*" Or arg型桁 Like "CLOB*" Then
    
        If is列名 Then
            editカラム値 = argカラム値
        Else
            editカラム値 = "'" & argカラム値 & "'"
        End If
        
    Else
        MsgBox "処理できない型：" & arg型桁
    End If

End Function

