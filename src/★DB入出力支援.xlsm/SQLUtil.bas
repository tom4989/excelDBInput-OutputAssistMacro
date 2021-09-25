Attribute VB_Name = "SQLUtil"
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

' 行情報
Private lngテーブル名記載行 As Long
Private lngカラム論理名記載行 As Long
Private lngカラム物理名記載行 As Long
Private lng型桁記載行 As Long
Private lngデータ開始行 As Long
Private lngデータ終了行 As Long

' 列情報
Private lngカラム終了列 As Long

' 名称
Private txtテーブル論理名, txtテーブル物理名 As String

' 状態
Private isHidden As Boolean
Private lngDBCount結果 As Long

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------

Public Property Set 隊粗油シート(ByRef arg対象シート)
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
' 機能：テーブル名記載行（＝試験データテーブルの開始位置）を設定する。
' *********************************************************************************************************************
'
Public Sub setテーブル名記載行(ByVal argテーブル名記載行)

    ' 行情報の設定
    lngテーブル名記載行 = argテーブル名記載行
    lngカラム論理名記載行 = argテーブル名記載行 + 1
    lngカラム物理名記載行 = argテーブル名記載行 + 2
    lng型桁記載行 = argテーブル名記載行 + 3
    lngデータ開始行 = argテーブル名記載行 + 4
    
    ' 列情報の設定
    lngカラム終了列 = ActiveSheet.Range("B" & lngカラム物理名記載行).End(xlToRight).Column
    
    ' 名称情報の設定
    txtテーブル論理名 = ActiveSheet.Range("A" & lngテーブル名記載行).Value
    txtテーブル物理名 = ActiveSheet.Range("D" & lngテーブル名記載行).Value

    ' 表示/非表示状態(論理名列の状態で判断)
    isHidden = ActiveSheet.Cells(lngカラム論理名記載行, 1).EntireRow.Hidden

End Sub

' *********************************************************************************************************************
' 機能：指定された列番号のカラムの論理名を返却する。
' *********************************************************************************************************************
'
Public Function getカラム論倫理(ByVal arg指定カラム列 As Long) As String

    getカラム論理名 = Me.対象シート.Cells(lngカラム論理名記載行, arg指定カラム列)

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
    
        addOrderBy = Re.Replace(txtQuery, "SELECT * FROM ( $1") & " ) "
    
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
        Dim txtSELECT文, txtWHERE句 As String
        
        For j = 2 To lngカラム終了列
        
            ' ---------------------------------------------------------------------------------------------------------
            ' SELECT句
            ' ---------------------------------------------------------------------------------------------------------
        
            If txtSELECT文 <> "SELECT " Then
                txtSELECT文 = txtSELECT文 & ", "
            End If
        
            txtSELECT文 = txtSELECT文 & _
                editカラム値(.Cells(lngカラム物理名記載行, j).Value, .Cells(lng型桁記載行, j).Value, True)
            
            ' ---------------------------------------------------------------------------------------------------------
            ' WHERE句
            ' ---------------------------------------------------------------------------------------------------------
        
            If lngデータ行 = -1 Then
                GoTo continue
            End If
        
            If .Cells(lngデータ行, j).Value <> "" Then
        
                If txtWHERE句 <> "" Then
                    txtWHERE句 = txtWHERE句 & " AND "
                Else
                    txtWHERE句 = " WHERE "
                End If
            
                txtWHERE句 = txtWHERE句 & _
                    .Cells(lngカラム物理名記載行, j).Value & " = " & _
                    editカラム値(.Cells(argデータ行, j).Value, .Cells(lng型桁記載行, j).Value, False)
            End If
continue:
        Next j
        
    End With

    createSELECT文From単行 = txtSELECT文 & " FROM " & txtテーブル物理名 & txtWHERE句

End Function

' *********************************************************************************************************************
' 機能：複数のデータ行をもとにSQL文を作成する。
' 　　　作成するSELECT文は、1テーブル（複数行）に対し、1SELECT文（複数のSELECT文をUNIONでまとめたもの）となる。
' *********************************************************************************************************************
'
Public Function createSELECT文From複数行() As String

    Dim txtQuery As String

    If getデータ行の入力数() > 0 Then
    
        Dim j As Long
        
        For j = lngデータ開始行 To lngデータ終了業
        
            If getデータ入力数(j) > 0 Then
            
                If txtQuery <> "" Then
                    
                    txtQuery = txtQuery & vbCrLf & " UNION "
                End If
            
                txtQuery = txtQuery & createSELECT文From単行(j)
            End If
             
         Next j
         
    Else
    
        txtQuery = txtQuery & createSELECT文From単行()
        
    End If
        
    ' -----------------------------------------------------------------------------------------------------------------
    ' ORDER B句の付与
    ' -----------------------------------------------------------------------------------------------------------------
        
    createSELECT文From複数行 = addOrderBy(txtQuery)
        
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
            
                txtQuery = addOrderBy(createSELECT文Fromt単行()) & ";" & vbCrLf
                Exit For ' 複数行選択されていても、同じSQLになるので1行だけ作成する
            End If
        Next k
    End If

    If strQuery <> "" Then
    
        strQuery = vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル論理名 & vbCrLf & txtQuery
        createSELECT文From複数行To複数SQL = txtQuery
        
    End If
        
End Function

' *********************************************************************************************************************
' 機能：INSERT文を作成する
' *********************************************************************************************************************
'
Public Function createInsert文(ByVal is選択行のみ As Boolean) As String

    Dim txt結果 As String
    
    With Me.対象シート
    
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
            
            Dim txtInsertInto As String: txtInsertInto = "INSERT INTO " & txtテーブル物理名 & " ("
            Dim txtInsertValues As String: txtInsertValues = " VALUES ("
            
            For k = 2 To lngカラム終了列
            
                If k > 2 Then
                    txtInsertInto = txtInsertInto & ", "
                    txtInsertValues = txtInsertValues & ", "
                End If
                
                txtInsertInto = txtInsertInto & .Cells(lngカラム物理名記載行, k)
                txtInsertValues = txtInsertValues & editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k))
            Next k
            
            txt結果 = txt結果 & txtInsertInto & ")" & vbCrLf & "   " & txtInsertValues & ");" & vbCrLf
            
jContinue:
        Next j
        
    End With

    If txt結果 <> "" Then
        txt結果 = vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル物理名 & vbCrLf & txt結果
        createInsert文 = txt結果
    End If

End Function

' *********************************************************************************************************************
' 機能：UPDATE文を作成する
' *********************************************************************************************************************
'
Public Function createUpdate文(ByVal is選択行のみ As Boolean) As String

    Dim txt結果 As String
    
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
            
            Dim txtUpdate As String: txtUpdate = "UPDATE " & txtテーブル物理名 & " SET "
            Dim txtWHERE As String: txtWHERE = " WHERE "
            
            For k = 2 To lngカラム終了列
            
                If containArray(var主キー, .Cells(lngカラム物理名記載行, k)) Then
                
                    If txtWHERE <> " WHERE " Then
                    
                        txtWHERE = txtWHERE & " AND "
                    End If
            
                    txtWHERE = txtWHERE & .Cells(lngカラム物理名記載行, k) _
                        & " = " & editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k))
                Else
                    If Not txtUpdate Like "* SET " Then
                    
                        txtUpdate = txtUpdate & " , "
                    End If
                    
                    txtUpdate = txtUpdate & .Cells(lngカラム物理名記載行, k) _
                        & " = " & editカラム値(.Cells(j, k), .Cells(lng型桁記載行, k))
                End If
                
            Next k
            
            txt結果 = txt結果 & txtUpdate & vbCrLf & "   " & txtWHERE & ";" & vbCrLf
            
jContinue:
        Next j
        
    End With

    If txt結果 <> "" Then
        txt結果 = vbCrLf & "-- " & txtテーブル論理名 & " " & txtテーブル物理名 & vbCrLf & txt結果
        createInsert文 = txt結果
    End If

End Function

' *********************************************************************************************************************
' 機能：SELECT文を、COUNTを行うSQLに変更する。
' *********************************************************************************************************************
'
Public Function createCount文(ByVal txtQuery As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "(SELECT .+ FROM)"
    
    createCount文 = Re.Replace(txtQuery, "SELECT COUNT(*) AS COUNT FROM")
    
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
    Me.対象シート.Rows(arg追加行番号).ClearContens ' 内容をクリア
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
' 機能：選択された行を網掛
' *********************************************************************************************************************
'
Public Sub edit選択行網掛(ByVal arg選択行番号 As Long)

    With Me.対象シート
    
        Call Me.edit網掛(.Range(.Cells(arg選択行番号, 2), .Cells(arg選択行番号, lngカラム終了列)))
        
    End With

End Sub

' *********************************************************************************************************************
' 機能：選択された範囲を網掛け
' *********************************************************************************************************************
'
Public Sub edit変更強調色(ByRef arg修飾範囲 As Range)

    With arg修飾範囲.Interior
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

    get主キー項目連結文字列 = Join(get主キー(lng対象データ行))

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
        arg対象データ行 = lngカラム論理名記載行
    End If

    Dim var主キー As Variant
    ReDim var主キー(1 To lngカラム終了列) ' 予想され得る最大値＝カラム数で配列を確保
    
    Dim i, lng主キー数 As Long
    
    For i = 2 To lngカラム終了列
    
        With Me.対象シート
        
            ' 主キーカラムであるか否かを、背景色で判断
            If .Cells(lngカラム物理名記載行, i).Interior.ThemeColor = xlThemeColorAccent2 Then
                
                lng主キー数 = lng主キー数 + 1
                var主キー(lng主キー数) = .Cells(arg対象データ行, i).Value
                
            End If
        End With
    Next
    
    If lng主キー数 = 0 Then
    
        get主キー = Empty
    Else
    
        ReDim Preserve var主キー(1 To lng主キー数)
        get主キー = var主キー
        
    End If
    
End Function


' *********************************************************************************************************************
' 機能：カラムに対する値を型桁に合わせて加工する(チェック機能付)
' *********************************************************************************************************************
'
Private Function editカラム値( _
    ByVal argカラム値 As String, ByVal arg型桁 As String, Optional ByVal is列名 = False) As String
    
    If argカラム値 = "" Then
        editカラム値 = "NULL"
        Exit Function
    End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    If argカラム値 Like "DATE*" Then
    
        If UCase(argカラム値) = "SYSTIMESTAMP" Or UCase(argカラム値) = "SYSDATE" Then
        
            editカラム値 = argカラム値
        Else
            If is列名 Then
                editカラム値 = "TO_CHAR(" & argカラム値 & ", 'YYYY/MM/DD HH24:MI:SS')"
            Else
                editカラム値 = "TO_DATE('" & argカラム値 & "', 'YYYY/MM/DD HH24:MI:SS')"
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
                editカラム値 = "TO_CHAR(" & argカラム値 & ", 'YYYY/MM/DD HH24:MI:SS.FF6')"
            Else
                editカラム値 = "TO_TIMESTAMP('" & argカラム値 & "', 'YYYY/MM/DD HH24:MI:SS.FF6')"
                
            End If
        End If

    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf argカラム値 Like "NUMBER*" Then
    
        editカラム値 = argカラム値

    ' -----------------------------------------------------------------------------------------------------------------
    ' VARCHAR2,CHAR,BLOB,CLOB型
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg型桁 Like "VARCHAR2*" Or arg型桁 Like "CHAR*" Or arg型桁 Like "BLOB*" Or arg型桁 Like "CLOB*" Then
    
        If is列名 Then
            editカラム値 = argカラム値
        Else
            editカラム値 = "'" & argカラム値 & "'"
        End If
        
    Else
        MsgBox "処理できない型：" & arg型桁
    End If

End Function

' *********************************************************************************************************************
' 機能：SQLを適度に整形
' *********************************************************************************************************************
'
Public Function SQL整形(ByVal txtSQL As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Global = True
    
    ' 正規化(改行や空白を統一する）
    Re.Pattern = "[\r\n]"
    txtSQL = Re.Replace(txtSQL, " ")
    
    Re.Pattern = " +"
    txtSQL = Re.Replace(txtSQL, " ")
    
    ' 前に改行
    Re.Pattern = "(AND) "
    txtSQL = Re.Replace(txtSQL, vbCrLf & "   $1")
    
    ' 前後に改行
    Re.Pattern = " (ORDER BY|WHERE|FROM|UNION) "
    txtSQL = Re.Replace(txtSQL, vbCrLf & "$1" & vbCrLf & "    ")
    
    ' 後方に開業
    Re.Pattern = "(SELECT) "
    txtSQL = Re.Replace(txtSQL, "$1" & vbCrLf & "    ")
    
    SQL整形 = txtSQL
    
End Function
