VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQL���� 
   Caption         =   "���݂̃V�[�g"
   ClientHeight    =   8388
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12840
   OleObjectBlob   =   "frmSQL����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSQL����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public wb�N�����u�b�N As Workbook
Public wb�O����s���� As Workbook

Private txtSQL�쐬���V�[�g�� As String
Private txt�g�����U�N�V�������ʕ����� As String
Private txtSQL���s�o�b�`�t�@�C���p�X As String
Private txtSQL���s���O�t�@�C���p�X As String
Private txtSQL�쐬���� As String
Private txt���ʃt�@�C���o�͐� As String

Private obj�ݒ�l�V�[�g As cls�ݒ�l�V�[�g

' *********************************************************************************************************************
' * �@�\�@�F�t�H�[������������
' *********************************************************************************************************************
'
Private Sub UserForm_Initialize()

    Set wb�N�����u�b�N = ActiveWorkbook
    
    btnSQL����.SetFocus

End Sub

' *********************************************************************************************************************
' * �@�\�@�F�t�H�[������������
' *********************************************************************************************************************
'
Public Sub �ݒ�l���[�h(arg�ݒ�l�V�[�g As cls�ݒ�l�V�[�g)

    Set obj�ݒ�l�V�[�g = arg�ݒ�l�V�[�g
    
    Dim txt�ڑ���� As Variant
    
    If Not obj�ݒ�l�V�[�g Is Nothing Then
        For Each txt�ڑ���� In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item("�ڑ����")
    
            cmb�ڑ����.AddItem (txt�ڑ����)
        
        Next txt�ڑ����
        
        cmb�ڑ����.ListIndex = 0
    
        txt���ʃt�@�C���o�͐� = Replace(obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item("���ʃt�@�C���o�͐�"), _
            "%USERPROFILE%", Environ("UserProfile"))

    End If
    
End Sub

' *********************************************************************************************************************
' * �@�\�@�F���R�[�h�擾�{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn���R�[�h�擾_Click()
        
    Call ���R�[�h�擾�{�^��
        
    Unload frmSQL����
        
End Sub

' *********************************************************************************************************************
' * �@�\�@�F���R�[�h�擾����
' *********************************************************************************************************************
'
Private Sub ���R�[�h�擾�{�^��()

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("���R�[�h�擾")
    
    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)
        
    ' �{����
        
    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    Call obj�����f�[�^�V�[�g.������(obj�ݒ�l�V�[�g, cmb�ڑ����.Text)
    
    Set wb�O����s���� = obj�����f�[�^�V�[�g.get���R�[�h(Nothing)
    
    If Not (wb�O����s���� Is Nothing) Then
        Application.CutCopyMode = False
        
        wb�O����s����.Activate
        wb�O����s����.ActiveSheet.Range("A1").Select
        
        btn���R�[�h�ǉ��擾.Enabled = True
        
    End If
        
    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("���R�[�h�擾")
        
   ' ����I��
   GoTo Finally
    
ErrorCatch:

Finally:

    txb�X�e�[�^�X�o�[.Value = get�ُ펞���b�Z�[�W("���R�[�h�擾")

    '���s�O�̏�Ԃɖ߂�
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)

End Sub

' *********************************************************************************************************************
' * �@�\�@�F�X�V�O���R�[�h�擾�{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn�X�V�O���R�[�h�擾_Click()

    ���R�[�h�擾�{�^��
    btn�X�V�ヌ�R�[�h�擾.Enabled = True

End Sub

' *********************************************************************************************************************
' * �@�\�@�F�X�V�ヌ�R�[�h�擾�{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn�X�V�ヌ�R�[�h�擾_Click()

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("�X�V�ヌ�R�[�h�擾")
    
    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)
        
    ' �{����
    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    
    Set wb�O����s���� = obj�����f�[�^�V�[�g.get���R�[�h(wb�O����s����)
    
    If Not (wb�O����s���� Is Nothing) Then
        Application.CutCopyMode = False
        
        wb�O����s����.Activate
        wb�O����s����.ActiveSheet.Range("A1").Select
        
        Call obj�����f�[�^�V�[�g.edit���s���ʍ���(wb�O����s����)
    
    End If

    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("�X�V�ヌ�R�[�h�擾")

    btn�X�V�ヌ�R�[�h�擾.Enabled = False

   ' ����I��
   GoTo Finally
    
ErrorCatch:

Finally:

    txb�X�e�[�^�X�o�[.Value = get�ُ펞���b�Z�[�W("���R�[�h�擾")

    '���s�O�̏�Ԃɖ߂�
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)

End Sub

' *********************************************************************************************************************
' * �@�\�@�FSQL�����{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btnSQL����_Click()

    '�G���[�g���b�v
    On Error GoTo ErrorCatch

    txtSQL�쐬���� = getTimestamp()

    txb�X�e�[�^�X�o�[.Value = get�J�n���b�Z�[�W("SQL����")

    Dim obj�A�v���P�[�V�������� As New cls�A�v���P�[�V��������
    obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (False)

    wb�N�����u�b�N.Activate
    
    Dim obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g
    Set obj�����f�[�^�V�[�g = New cls�����f�[�^�V�[�g
    Call obj�����f�[�^�V�[�g.������(obj�ݒ�l�V�[�g, cmb�ڑ����.Text)
    
    Dim obj�ΏۃV�[�g As Worksheet
    Set obj�ΏۃV�[�g = ActiveSheet
    txtSQL�쐬���V�[�g�� = ActiveSheet.Name
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    txt�g�����U�N�V�������ʕ����� = FSO.getBaseName(wb�N�����u�b�N.Name) & "_" & txtSQL�쐬���V�[�g�� & "_" & getTimestamp
    
    txtSQL.Value = ""
    frmSQL����.Repaint
    
    ' SPOOL���̏o��
    If ckbSPOOL Then
        txtSQL.Value = txtSQL.Value & vbCrLf & _
            "SPOOL """ & txt�g�����U�N�V�������ʕ����� & ".sql,log" & vbCrLf & vbCrLf
    End If
    
    If True Then
    
        txtSQL.Value = vbCrLf & "-- �V�[�g���F" & obj�ΏۃV�[�g.Name & vbLf & _
            createSQL��(obj�����f�[�^�V�[�g, obj�ΏۃV�[�g)
    Else
    
        Dim strSQL As String
        
        For Each obj�ΏۃV�[�g In ActiveWorkbook.Sheets
        
            obj�ΏۃV�[�g.Activate
            
            strSQL = strSQL & vbLf & "-- �V�[�g���F" & obj�ΏۃV�[�g.Name
            strSQL = strSQL & createSQL��(obj�����f�[�^�V�[�g, obj�ΏۃV�[�g)
            
        Next
        
        txtSQL.Value = strSQL
    
    End If
    
    ' �㑱�̃{�^����enable�ɕύX
    btn�N���b�v�{�[�h�ɃR�s�[.Enabled = True
    btn���ʂ��t�@�C���ɏo��.Enabled = True
    btn�o�̓t�@�C�������s.Enabled = False
    btn�G���[�m�F.Enabled = False


    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("SQL����")
    
   ' ����I��
   GoTo Finally
    
ErrorCatch:

    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("SQL����")

Finally:

   '���s�O�̏�Ԃɖ߂�
   obj�A�v���P�[�V��������.�A�v���P�[�V��������ؑ� (True)
        
End Sub

' *********************************************************************************************************************
' * �@�\�@�FSQL����������
' *********************************************************************************************************************
'
Private Function createSQL��( _
    ByRef obj�����f�[�^�V�[�g As cls�����f�[�^�V�[�g, _
    ByRef obj�ΏۃV�[�g As Worksheet) As String
    
    If rdInsert Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.INSERT��, rdb�I���s�̂�)
        
    ElseIf rdUpdate Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.UPDATE��, rdb�I���s�̂�)
        
    ElseIf rdSELECT Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.SELECT��, rdb�I���s�̂�)
        
    ElseIf rdDelete Then
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.DELETE��, rdb�I���s�̂�)
        
    ElseIf rdDELETEINSERT Then
        
        createSQL�� = obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.DELETE��, rdb�I���s�̂�) & _
            obj�����f�[�^�V�[�g.�ΏۃV�[�gSQL���쐬(obj�ΏۃV�[�g, SQL���.INSERT��, rdb�I���s�̂�)
        
    End If
    
End Function

' *********************************************************************************************************************
' * �@�\�@�F�N���b�v�{�[�h�ɃR�s�[
' *********************************************************************************************************************
'
Private Sub btn�N���b�v�{�[�h�ɃR�s�[_Click()

    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = txtSQL.Value
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
    
    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("�N���b�v�{�[�h�ɃR�s�[")

End Sub

' *********************************************************************************************************************
' * �@�\�@�F���ʂ��t�@�C���ɏo�̓{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn���ʂ��t�@�C���ɏo��_Click()

    mkdirIFNotExist txt���ʃt�@�C���o�͐�

    Dim txt�o�̓p�X As String
    txt�o�̓p�X = txt���ʃt�@�C���o�͐� & "\" & txtSQL�쐬���V�[�g�� & "_" & txtSQL�쐬���� & ".sql"

    Open txt�o�̓p�X For Output As #1
    
    Dim txtSQL�t�@�C�����e As String
    txtSQL�t�@�C�����e = "SPOOL " & txtSQL�쐬���V�[�g�� & "_" & txtSQL�쐬���� & ".log" & vbCrLf & _
        txtSQL.Value & vbCrLf & _
        "SPOOL OFF" & vbCrLf & _
        "EXIT"
    
    Print #1, txtSQL�t�@�C�����e
    Close #1

    Dim txt�o�b�`�t�@�C�����e As String
    
    txt�o�b�`�t�@�C�����e = "cd " & txt���ʃt�@�C���o�͐� & vbCrLf
    
    If obj�ݒ�l�V�[�g.�ݒ�l���X�g.Exists("ORACLE_HOME") Then
     
        txt�o�b�`�t�@�C�����e = txt�o�b�`�t�@�C�����e & _
            "set ORACLE_HOME=" & obj�ݒ�l�V�[�g.�ݒ�l���X�g("ORACLE_HOME") & vbCrLf
    End If
    
    txt�o�b�`�t�@�C�����e = txt�o�b�`�t�@�C�����e & _
        "sqlplus " & _
        obj�ݒ�l�V�[�g.�ݒ�l���X�g("�ڑ����").Item(cmb�ڑ����.Value).Item("UID") & "/" & _
        obj�ݒ�l�V�[�g.�ݒ�l���X�g("�ڑ����").Item(cmb�ڑ����.Value).Item("PWD") & "@" & _
        obj�ݒ�l�V�[�g.�ݒ�l���X�g("�ڑ����").Item(cmb�ڑ����.Value).Item("DSN") & _
        " @" & txt�o�̓p�X & vbCrLf & "pause"

    Open txt�o�̓p�X & ".bat" For Output As #1
    Print #1, txt�o�b�`�t�@�C�����e
    Close #1

    ' ���ʏo��
    txb�X�e�[�^�X�o�[.Value = get�I�����b�Z�[�W("���ʂ��t�@�C���ɏo��")
    txb�X�e�[�^�X�o�[ = txb�X�e�[�^�X�o�[.Value & vbCr & txt�o�̓p�X

    btn�o�̓t�@�C�������s.Enabled = True

End Sub

' *********************************************************************************************************************
' * �@�\�@�F���ʂ��t�@�C���ɏo�̓{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn�o�̓t�@�C�������s_Click()

    Dim txt�o�̓p�X As String
    txt�o�̓p�X = txt���ʃt�@�C���o�͐� & "\" & txtSQL�쐬���V�[�g�� & "_" & txtSQL�쐬���� & ".sql.bat"

    Call Shell(txt�o�̓p�X, vbNormalFocus)

    btn�G���[�m�F.Enabled = True
    
End Sub

' *********************************************************************************************************************
' * �@�\�@�F�G���[�m�F�{�^���������̏���
' *********************************************************************************************************************
'
Private Sub btn�G���[�m�F_Click()

    Dim txt�o�̓p�X As String
    txt�o�̓p�X = txt���ʃt�@�C���o�͐� & "\" & txtSQL�쐬���V�[�g�� & "_" & txtSQL�쐬���� & ".log"

    Call Shell("notepad.exe " & txt�o�̓p�X, vbNormalFocus)

End Sub
