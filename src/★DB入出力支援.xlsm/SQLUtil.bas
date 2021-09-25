Attribute VB_Name = "SQLUtil"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

Const cnst�C���f���g = "   "

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

' �ΏۃV�[�g
Private obj�ΏۃV�[�g As Worksheet

' �s���
Private lng�e�[�u�����L�ڍs As Long
Private lng�J�����_�����L�ڍs As Long
Private lng�J�����������L�ڍs As Long
Private lng�^���L�ڍs As Long
Private lng�f�[�^�J�n�s As Long
Private lng�f�[�^�I���s As Long

' ����
Private lng�J�����I���� As Long

' ����
Private txt�e�[�u���_����, txt�e�[�u�������� As String

' ���
Private isHidden As Boolean
Private lngDBCount���� As Long

' ---------------------------------------------------------------------------------------------------------------------
' Property
' ---------------------------------------------------------------------------------------------------------------------

Public Property Set ���e���V�[�g(ByRef arg�ΏۃV�[�g)
    Set obj�ΏۃV�[�g = arg�ΏۃV�[�g
End Property

Public Property Get �ΏۃV�[�g() As Worksheet

    If obj�ΏۃV�[�g Is Nothing Then
        Set obj�ΏۃV�[�g = ActiveSheet
    End If
    
    Set �ΏۃV�[�g = obj�ΏۃV�[�g
    
End Property

Public Property Get �J�����_�����L�ڍs()
    �J�����_�����L�ڍs = lng�J�����_�����L�ڍs
End Property

Public Property Get �f�[�^�J�n�s()
    �f�[�^�J�n�s = lng�f�[�^�J�n�s
End Property

Public Property Get �f�[�^�I���s()
    �f�[�^�I���s = lng�f�[�^�I���s
End Property

Public Property Get �J�����I����()
    �J�����I���� = lng�J�����I����
End Property

' *********************************************************************************************************************
' �@�\�F�e�[�u�����L�ڍs�i�������f�[�^�e�[�u���̊J�n�ʒu�j��ݒ肷��B
' *********************************************************************************************************************
'
Public Sub set�e�[�u�����L�ڍs(ByVal arg�e�[�u�����L�ڍs)

    ' �s���̐ݒ�
    lng�e�[�u�����L�ڍs = arg�e�[�u�����L�ڍs
    lng�J�����_�����L�ڍs = arg�e�[�u�����L�ڍs + 1
    lng�J�����������L�ڍs = arg�e�[�u�����L�ڍs + 2
    lng�^���L�ڍs = arg�e�[�u�����L�ڍs + 3
    lng�f�[�^�J�n�s = arg�e�[�u�����L�ڍs + 4
    
    ' ����̐ݒ�
    lng�J�����I���� = ActiveSheet.Range("B" & lng�J�����������L�ڍs).End(xlToRight).Column
    
    ' ���̏��̐ݒ�
    txt�e�[�u���_���� = ActiveSheet.Range("A" & lng�e�[�u�����L�ڍs).Value
    txt�e�[�u�������� = ActiveSheet.Range("D" & lng�e�[�u�����L�ڍs).Value

    ' �\��/��\�����(�_������̏�ԂŔ��f)
    isHidden = ActiveSheet.Cells(lng�J�����_�����L�ڍs, 1).EntireRow.Hidden

End Sub

' *********************************************************************************************************************
' �@�\�F�w�肳�ꂽ��ԍ��̃J�����̘_������ԋp����B
' *********************************************************************************************************************
'
Public Function get�J�����_�ϗ�(ByVal arg�w��J������ As Long) As String

    get�J�����_���� = Me.�ΏۃV�[�g.Cells(lng�J�����_�����L�ڍs, arg�w��J������)

End Function


' *********************************************************************************************************************
' �@�\�F�Z���ɓ��͂���Ă���f�[�^������ԋp����B
' *********************************************************************************************************************
'
Public Function get����()

    get���� = Me.�ΏۃV�[�g.Cells(lng�e�[�u�����L�ڍs, 7)

End Function

' *********************************************************************************************************************
' �@�\�FSELECT���ɁAORDER BY��t�^����B
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

    Dim var��L�[ As Variant
    var��L�[ = get��L�[()
    
    If Not IsEmpty(var��L�[) Then
        addOrderBy = addOrderBy & " ORDER BY " & Join(var��L�[, ", ")
    End If

End Function



' *********************************************************************************************************************
' �@�\�F1�s�̃f�[�^�s�����Ƃ�SQL�����쐬����B
' �@�@�@�f�[�^�s���w�肵�Ȃ��ꍇ�AWHERE��Ȃ���SELECT�����쐬����B
' *********************************************************************************************************************
'
Public Function createSELECT��From�P�s(Optional arg�f�[�^�s As Long = -1) As String

    With ActiveSheet
        
        Dim j As Long
        Dim txtSELECT��, txtWHERE�� As String
        
        For j = 2 To lng�J�����I����
        
            ' ---------------------------------------------------------------------------------------------------------
            ' SELECT��
            ' ---------------------------------------------------------------------------------------------------------
        
            If txtSELECT�� <> "SELECT " Then
                txtSELECT�� = txtSELECT�� & ", "
            End If
        
            txtSELECT�� = txtSELECT�� & _
                edit�J�����l(.Cells(lng�J�����������L�ڍs, j).Value, .Cells(lng�^���L�ڍs, j).Value, True)
            
            ' ---------------------------------------------------------------------------------------------------------
            ' WHERE��
            ' ---------------------------------------------------------------------------------------------------------
        
            If lng�f�[�^�s = -1 Then
                GoTo continue
            End If
        
            If .Cells(lng�f�[�^�s, j).Value <> "" Then
        
                If txtWHERE�� <> "" Then
                    txtWHERE�� = txtWHERE�� & " AND "
                Else
                    txtWHERE�� = " WHERE "
                End If
            
                txtWHERE�� = txtWHERE�� & _
                    .Cells(lng�J�����������L�ڍs, j).Value & " = " & _
                    edit�J�����l(.Cells(arg�f�[�^�s, j).Value, .Cells(lng�^���L�ڍs, j).Value, False)
            End If
continue:
        Next j
        
    End With

    createSELECT��From�P�s = txtSELECT�� & " FROM " & txt�e�[�u�������� & txtWHERE��

End Function

' *********************************************************************************************************************
' �@�\�F�����̃f�[�^�s�����Ƃ�SQL�����쐬����B
' �@�@�@�쐬����SELECT���́A1�e�[�u���i�����s�j�ɑ΂��A1SELECT���i������SELECT����UNION�ł܂Ƃ߂����́j�ƂȂ�B
' *********************************************************************************************************************
'
Public Function createSELECT��From�����s() As String

    Dim txtQuery As String

    If get�f�[�^�s�̓��͐�() > 0 Then
    
        Dim j As Long
        
        For j = lng�f�[�^�J�n�s To lng�f�[�^�I����
        
            If get�f�[�^���͐�(j) > 0 Then
            
                If txtQuery <> "" Then
                    
                    txtQuery = txtQuery & vbCrLf & " UNION "
                End If
            
                txtQuery = txtQuery & createSELECT��From�P�s(j)
            End If
             
         Next j
         
    Else
    
        txtQuery = txtQuery & createSELECT��From�P�s()
        
    End If
        
    ' -----------------------------------------------------------------------------------------------------------------
    ' ORDER B��̕t�^
    ' -----------------------------------------------------------------------------------------------------------------
        
    createSELECT��From�����s = addOrderBy(txtQuery)
        
End Function


' *********************************************************************************************************************
' �@�\�F�����̃f�[�^�s������SQL�����쐬����B�쐬����SELECT���́A1�f�[�^�ɑ΂��A1SELECT���ƂȂ�B
' *********************************************************************************************************************
'
Public Function createSELECT��From�����sTo����SQL(Optional ByVal is�I���s�̂� As Boolean = False) As String

    Dim txtQuery As String
    
    ' �f�[�^�s�̂����ꂩ�ɉ���������͂���Ă���ꍇ
    If get�f�[�^�s�̓��͐�() > 0 Then
    
        Dim j As Long

        For j = lng�f�[�^�J�n�s To lng�f�[�^�I���s
        
            If is�I���s�̂� And Not is�I�����(j) Then ' �I���s�̂�SQL���쐬�Ώۂɂ���ꍇ�̍l��
            
                GoTo jContinue
            End If
            
            If get�f�[�^�s�̓��͐�(j) > 0 Then
            
                txtQuery = txtQuery & addOrderBy(createSELECT��From�P�s(j)) & ";" & vbCrLf
            
            End If
jContinue:

        Next j

    End If

    ' �O������SQL���쐬����Ă��Ȃ��ꍇ
    
    If txtQuery = "" Then
    
        Dim k As Long
        
        For k = lng�f�[�^�J�n�s To lng�f�[�^�I���s
        
            If is�I���s�̂� And Not is�I�����(k) Then ' �I���s�̂�SQL���쐬�Ώۂɂ���ꍇ�̍l��
                ' �������Ȃ�
            Else
            
                txtQuery = addOrderBy(createSELECT��Fromt�P�s()) & ";" & vbCrLf
                Exit For ' �����s�I������Ă��Ă��A����SQL�ɂȂ�̂�1�s�����쐬����
            End If
        Next k
    End If

    If strQuery <> "" Then
    
        strQuery = vbCrLf & "-- " & txt�e�[�u���_���� & " " & txt�e�[�u���_���� & vbCrLf & txtQuery
        createSELECT��From�����sTo����SQL = txtQuery
        
    End If
        
End Function

' *********************************************************************************************************************
' �@�\�FINSERT�����쐬����
' *********************************************************************************************************************
'
Public Function createInsert��(ByVal is�I���s�̂� As Boolean) As String

    Dim txt���� As String
    
    With Me.�ΏۃV�[�g
    
        If .Range("B" & lng�f�[�^�J�n�s).Value = "" Then
        
            Exit Function ' �f�[�^���Ȃ��ꍇ�AINSERT���̍쐬�͂��Ȃ�
        End If
        
        Dim j, k As Long
        
        ' �f�[�^�s�A1�s���Ƃ̏���
        For j = lng�f�[�^�J�n�s To lng�f�[�^�I���s
        
            If .Range("B" & j).Value = "" Then
            
                GoTo jContinue
            End If
            
            If is�I���s�̂� And Not is�I�����(j) Then ' �I���s�̂�SQL���쐬�Ώۂɂ���ꍇ�̍l��
            
                GoTo jContinue
            End If
            
            Dim txtInsertInto As String: txtInsertInto = "INSERT INTO " & txt�e�[�u�������� & " ("
            Dim txtInsertValues As String: txtInsertValues = " VALUES ("
            
            For k = 2 To lng�J�����I����
            
                If k > 2 Then
                    txtInsertInto = txtInsertInto & ", "
                    txtInsertValues = txtInsertValues & ", "
                End If
                
                txtInsertInto = txtInsertInto & .Cells(lng�J�����������L�ڍs, k)
                txtInsertValues = txtInsertValues & edit�J�����l(.Cells(j, k), .Cells(lng�^���L�ڍs, k))
            Next k
            
            txt���� = txt���� & txtInsertInto & ")" & vbCrLf & "   " & txtInsertValues & ");" & vbCrLf
            
jContinue:
        Next j
        
    End With

    If txt���� <> "" Then
        txt���� = vbCrLf & "-- " & txt�e�[�u���_���� & " " & txt�e�[�u�������� & vbCrLf & txt����
        createInsert�� = txt����
    End If

End Function

' *********************************************************************************************************************
' �@�\�FUPDATE�����쐬����
' *********************************************************************************************************************
'
Public Function createUpdate��(ByVal is�I���s�̂� As Boolean) As String

    Dim txt���� As String
    
    With Me.�ΏۃV�[�g
    
        If .Range("B" & lng�f�[�^�J�n�s).Value = "" Then
        
            Exit Function ' �f�[�^���Ȃ��ꍇ�AUPDATE���̍쐬�͂��Ȃ�
        End If
        
        Dim var��L�[ As Variant
        var��L�[ = Me.get��L�[()
        
        Dim j, k As Long
        
        ' �f�[�^�s�A1�s���Ƃ̏���
        For j = lng�f�[�^�J�n�s To lng�f�[�^�I���s
        
            If .Range("B" & j).Value = "" Then
            
                GoTo jContinue
            End If
            
            If is�I���s�̂� And Not is�I�����(j) Then ' �I���s�̂�SQL���쐬�Ώۂɂ���ꍇ�̍l��
            
                GoTo jContinue
            End If
            
            Dim txtUpdate As String: txtUpdate = "UPDATE " & txt�e�[�u�������� & " SET "
            Dim txtWHERE As String: txtWHERE = " WHERE "
            
            For k = 2 To lng�J�����I����
            
                If containArray(var��L�[, .Cells(lng�J�����������L�ڍs, k)) Then
                
                    If txtWHERE <> " WHERE " Then
                    
                        txtWHERE = txtWHERE & " AND "
                    End If
            
                    txtWHERE = txtWHERE & .Cells(lng�J�����������L�ڍs, k) _
                        & " = " & edit�J�����l(.Cells(j, k), .Cells(lng�^���L�ڍs, k))
                Else
                    If Not txtUpdate Like "* SET " Then
                    
                        txtUpdate = txtUpdate & " , "
                    End If
                    
                    txtUpdate = txtUpdate & .Cells(lng�J�����������L�ڍs, k) _
                        & " = " & edit�J�����l(.Cells(j, k), .Cells(lng�^���L�ڍs, k))
                End If
                
            Next k
            
            txt���� = txt���� & txtUpdate & vbCrLf & "   " & txtWHERE & ";" & vbCrLf
            
jContinue:
        Next j
        
    End With

    If txt���� <> "" Then
        txt���� = vbCrLf & "-- " & txt�e�[�u���_���� & " " & txt�e�[�u�������� & vbCrLf & txt����
        createInsert�� = txt����
    End If

End Function

' *********************************************************************************************************************
' �@�\�FSELECT�����ACOUNT���s��SQL�ɕύX����B
' *********************************************************************************************************************
'
Public Function createCount��(ByVal txtQuery As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "(SELECT .+ FROM)"
    
    createCount�� = Re.Replace(txtQuery, "SELECT COUNT(*) AS COUNT FROM")
    
End Function
' *********************************************************************************************************************
' �@�\�F�g���N���b�v�{�[�h�ɃR�s�[����
' *********************************************************************************************************************
'
Public Sub copy�gTo�N���b�v�{�[�h()

    ActiveWorkbook.ActiveSheet.Rows(lng�e�[�u�����L�ڍs & ":" & lng�f�[�^�J�n�s).Copy
    
End Sub


' *********************************************************************************************************************
' �@�\�F�f�[�^�s�̓��e���N���A����
' *********************************************************************************************************************
'
Public Sub clear�f�[�^�s()

    Me.�ΏۃV�[�g.Rows(lng�f�[�^�J�n�s & ";" & lng�f�[�^�I���s).ClearContents ' ���e���N���A
    Me.�ΏۃV�[�g.Rows(lng�f�[�^�J�n�s & ";" & lng�f�[�^�I���s).ClearComments ' �R�����g���N���A
    
End Sub

' *********************************************************************************************************************
' �@�\�F�����Ŏw�肳�ꂽ�����ݒ�l�Ɠ���̃J���������e�[�u�����ɑ��݂���ꍇ�A�ݒ�l��1�s�ڂɃZ�b�g����B
' *********************************************************************************************************************
'
Public Sub set���o����(ByVal var�����ݒ�l As Variant)

    With Me.�ΏۃV�[�g
    
        Dim i As Long
        For i = 2 To lng�J�����I����
    
            Dim j As Long
            For j = LBound(var�����ݒ�l) To UBound(var�����ݒ�l)
        
                If .Cells(lng�J�����������L�ڍs, i) = var�����ݒ�l(j, 1) Then
            
                    .Cells(lng�J�����������L�ڍs, i) = var�����ݒ�l(j, 2)
                
                End If
            Next j
        Next i
    End With
End Sub

' *********************************************************************************************************************
' �@�\�F�w�肳�ꂽ�s�ɋ�s���쐬����
' *********************************************************************************************************************
'
Public Sub add��s(ByVal arg�ǉ��s�ԍ� As Long)

    Me.�ΏۃV�[�g.Rows(lng�f�[�^�J�n�s).Copy  ' �f�[�^�s��1�s�ڂ��珑���R�s�[
    Me.�ΏۃV�[�g.Rows(arg�ǉ��s�ԍ�).Insert ' �s�ǉ�
    Me.�ΏۃV�[�g.Rows(arg�ǉ��s�ԍ�).ClearContens ' ���e���N���A
    Me.�ΏۃV�[�g.Rows(arg�ǉ��s�ԍ�).ClearComments ' �R�����g���N���A
    
    Application.CutCopyMode = False
    
End Sub


' *********************************************************************************************************************
' �@�\�F�I�����ꂽ�s�����F�Œ��F
' *********************************************************************************************************************
'
Public Sub edit�I���s����(ByVal arg�I���s�ԍ� As Long)

    With Me.�ΏۃV�[�g
    
        Call Me.edit�ύX�����F(.Range(.Cells(arg�I���s�ԍ�, 2), .Cells(arg�I���s�ԍ�, lng�J�����I����)))

    End With

End Sub

' *********************************************************************************************************************
' �@�\�F�I�����ꂽ�s��Ԋ|
' *********************************************************************************************************************
'
Public Sub edit�I���s�Ԋ|(ByVal arg�I���s�ԍ� As Long)

    With Me.�ΏۃV�[�g
    
        Call Me.edit�Ԋ|(.Range(.Cells(arg�I���s�ԍ�, 2), .Cells(arg�I���s�ԍ�, lng�J�����I����)))
        
    End With

End Sub

' *********************************************************************************************************************
' �@�\�F�I�����ꂽ�͈͂�Ԋ|��
' *********************************************************************************************************************
'
Public Sub edit�ύX�����F(ByRef arg�C���͈� As Range)

    With arg�C���͈�.Interior
        .Pattern = xlGray16
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

' *********************************************************************************************************************
' �@�\�F�w�肳�ꂽ��L�[���ڂ�A�������������ԋp����
' *********************************************************************************************************************
'
Function get��L�[���ژA��������(ByVal arg�Ώۃf�[�^�s As Long) As String

    get��L�[���ژA�������� = Join(get��L�[(lng�Ώۃf�[�^�s))

End Function

' *********************************************************************************************************************
' �@�\�F�f�[�^�s�̓��͐���ԋp����B
' *********************************************************************************************************************
'
Private Function get�f�[�^�s�̓��͐�(Optional arg�Ώۃf�[�^�s = -1)

    With ActiveSheet
    
        If arg�Ώۃf�[�^�s = -1 Then
        
            get�f�[�^�s�̓��͐� = WorksheetFunction.CountA( _
                .Range(.Cells(lng�f�[�^�J�n�s, 2), .Cells(lng�f�[�^�I���s, lng�J�����I����)))
            
        Else
            
            get�f�[�^�s�̓��͐� = WorksheetFunction.CountA( _
                .Range(.Cells(arg�Ώۃf�[�^�s, 2), .Cells(arg�Ώۃf�[�^�s, lng�J�����I����)))
                
        End If
            
    End With

End Function

' *********************************************************************************************************************
' �@�\�F��L�[��z��ŕԋp����
' *********************************************************************************************************************
'
Public Function get��L�[(Optional ByVal arg�Ώۃf�[�^�s As Long = -1) As Variant

    If arg�Ώۃf�[�^�s = -1 Then
        arg�Ώۃf�[�^�s = lng�J�����_�����L�ڍs
    End If

    Dim var��L�[ As Variant
    ReDim var��L�[(1 To lng�J�����I����) ' �\�z���꓾��ő�l���J�������Ŕz����m��
    
    Dim i, lng��L�[�� As Long
    
    For i = 2 To lng�J�����I����
    
        With Me.�ΏۃV�[�g
        
            ' ��L�[�J�����ł��邩�ۂ����A�w�i�F�Ŕ��f
            If .Cells(lng�J�����������L�ڍs, i).Interior.ThemeColor = xlThemeColorAccent2 Then
                
                lng��L�[�� = lng��L�[�� + 1
                var��L�[(lng��L�[��) = .Cells(arg�Ώۃf�[�^�s, i).Value
                
            End If
        End With
    Next
    
    If lng��L�[�� = 0 Then
    
        get��L�[ = Empty
    Else
    
        ReDim Preserve var��L�[(1 To lng��L�[��)
        get��L�[ = var��L�[
        
    End If
    
End Function


' *********************************************************************************************************************
' �@�\�F�J�����ɑ΂���l���^���ɍ��킹�ĉ��H����(�`�F�b�N�@�\�t)
' *********************************************************************************************************************
'
Private Function edit�J�����l( _
    ByVal arg�J�����l As String, ByVal arg�^�� As String, Optional ByVal is�� = False) As String
    
    If arg�J�����l = "" Then
        edit�J�����l = "NULL"
        Exit Function
    End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE�^
    ' -----------------------------------------------------------------------------------------------------------------
    '
    If arg�J�����l Like "DATE*" Then
    
        If UCase(arg�J�����l) = "SYSTIMESTAMP" Or UCase(arg�J�����l) = "SYSDATE" Then
        
            edit�J�����l = arg�J�����l
        Else
            If is�� Then
                edit�J�����l = "TO_CHAR(" & arg�J�����l & ", 'YYYY/MM/DD HH24:MI:SS')"
            Else
                edit�J�����l = "TO_DATE('" & arg�J�����l & "', 'YYYY/MM/DD HH24:MI:SS')"
            End If
        End If
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE�^
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg�^�� Like "TIMESTAMP*" Then
        
        If UCase(arg�J�����l) = "SYSTIMESTAMP" Or UCase(arg�J�����l) = "SYSDATE" Then
        
            edit�J�����l = arg�J�����l
        Else
            If is�� Then
                edit�J�����l = "TO_CHAR(" & arg�J�����l & ", 'YYYY/MM/DD HH24:MI:SS.FF6')"
            Else
                edit�J�����l = "TO_TIMESTAMP('" & arg�J�����l & "', 'YYYY/MM/DD HH24:MI:SS.FF6')"
                
            End If
        End If

    ' -----------------------------------------------------------------------------------------------------------------
    ' DATE�^
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg�J�����l Like "NUMBER*" Then
    
        edit�J�����l = arg�J�����l

    ' -----------------------------------------------------------------------------------------------------------------
    ' VARCHAR2,CHAR,BLOB,CLOB�^
    ' -----------------------------------------------------------------------------------------------------------------
    '
    ElseIf arg�^�� Like "VARCHAR2*" Or arg�^�� Like "CHAR*" Or arg�^�� Like "BLOB*" Or arg�^�� Like "CLOB*" Then
    
        If is�� Then
            edit�J�����l = arg�J�����l
        Else
            edit�J�����l = "'" & arg�J�����l & "'"
        End If
        
    Else
        MsgBox "�����ł��Ȃ��^�F" & arg�^��
    End If

End Function

' *********************************************************************************************************************
' �@�\�FSQL��K�x�ɐ��`
' *********************************************************************************************************************
'
Public Function SQL���`(ByVal txtSQL As String) As String

    Dim Re As Object: Set Re = CreateObject("VBScript.RegExp")
    Re.Global = True
    
    ' ���K��(���s��󔒂𓝈ꂷ��j
    Re.Pattern = "[\r\n]"
    txtSQL = Re.Replace(txtSQL, " ")
    
    Re.Pattern = " +"
    txtSQL = Re.Replace(txtSQL, " ")
    
    ' �O�ɉ��s
    Re.Pattern = "(AND) "
    txtSQL = Re.Replace(txtSQL, vbCrLf & "   $1")
    
    ' �O��ɉ��s
    Re.Pattern = " (ORDER BY|WHERE|FROM|UNION) "
    txtSQL = Re.Replace(txtSQL, vbCrLf & "$1" & vbCrLf & "    ")
    
    ' ����ɊJ��
    Re.Pattern = "(SELECT) "
    txtSQL = Re.Replace(txtSQL, "$1" & vbCrLf & "    ")
    
    SQL���` = txtSQL
    
End Function
